[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$OutputDirectory,
    [string]$Mode = 'Verify',
    [switch]$SkipManual
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $projectRoot 'CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $projectRoot 'CreateOrder.xlsm.modules' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $projectRoot 'CreateOrderReleases' }
$WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
$SourceDirectory = [IO.Path]::GetFullPath($SourceDirectory)
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$reportDirectory = Join-Path $projectRoot "Trash\release-gate-$stamp"
$logDirectory = Join-Path $reportDirectory 'logs'
$reportJsonPath = Join-Path $reportDirectory 'report.json'
$reportMarkdownPath = Join-Path $reportDirectory 'report.md'
$backupPath = $null
$releasePath = $null
$gateResults = [Collections.Generic.List[object]]::new()
$overallCode = 0
$firstFailure = $null

New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null

function Add-GateResult {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Status,
        [int]$ExitCode = 0,
        [long]$DurationMs = 0,
        [string]$Command = '',
        [string]$Message = '',
        [string]$LogPath = ''
    )
    $result = [pscustomobject]@{
        id = $Id
        status = $Status
        duration_ms = $DurationMs
        command = $Command
        exit_code = $ExitCode
        message = $Message
        log = $LogPath
    }
    [void]$gateResults.Add($result)
    if ($Status -eq 'FAIL' -and $null -eq $script:firstFailure) { $script:firstFailure = $result }
    $colour = if ($Status -eq 'PASS') { 'Green' } elseif ($Status -eq 'WARN' -or $Status -eq 'MANUAL_REQUIRED') { 'Yellow' } else { 'Red' }
    Write-Host ("[{0}] {1} ({2} ms) {3}" -f $Status, $Id, $DurationMs, $Message) -ForegroundColor $colour
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

function Normalize-VbaCode {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    $normalized = $Text -replace "`r`n?", "`n"
    $lines = $normalized -split "`n"
    $start = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*Option\s+(Explicit|Compare)\b') { $start = $i; break }
    }
    $body = if ($start -gt 0) { $lines[$start..($lines.Count - 1)] } else { $lines }
    $result = (($body | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' -and $_ -notmatch '^\s*Attribute\s+m\w+\.VB_' -and $_ -notmatch "^\s*'" } | ForEach-Object { $_.TrimEnd() }) -join "`n").Trim()
    try {
        $repaired = ([Text.UTF8Encoding]::new($false, $true)).GetString([Text.Encoding]::GetEncoding(1251).GetBytes($result))
        if ($repaired -match '[А-Яа-яЁё]') { $result = $repaired }
    }
    catch {}
    return $result
}

function Get-SourceCodeVariants {
    param([Parameter(Mandatory = $true)][string]$Path)
    $bytes = [IO.File]::ReadAllBytes($Path)
    $variants = [Collections.Generic.List[string]]::new()
    try {
        $utf8 = [Text.UTF8Encoding]::new($false, $true)
        [void]$variants.Add((Normalize-VbaCode ($utf8.GetString($bytes))))
    }
    catch {
    }
    $cp1251 = [Text.Encoding]::GetEncoding(1251).GetString($bytes)
    $cpVariant = Normalize-VbaCode $cp1251
    if ($variants -notcontains $cpVariant) {
        [void]$variants.Add($cpVariant)
    }
    return @($variants)
}

function Test-SourceBookSynchronization {
    param([Parameter(Mandatory = $true)][string]$BookPath)
    $excel = $null
    $workbook = $null
    $mismatches = [Collections.Generic.List[string]]::new()
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AutomationSecurity = 1 } catch {}
        $workbook = $excel.Workbooks.Open($BookPath, 0, $true)
        foreach ($sourceFile in (Get-ChildItem -LiteralPath $SourceDirectory -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm' } | Sort-Object Name)) {
            $component = $null
            try { $component = $workbook.VBProject.VBComponents.Item($sourceFile.BaseName) } catch {
                [void]$mismatches.Add("missing component $($sourceFile.BaseName)")
                continue
            }
            try {
                $module = $component.CodeModule
                $embedded = if ($module.CountOfLines -gt 0) { $module.Lines(1, $module.CountOfLines) } else { '' }
                $embeddedNormalized = Normalize-VbaCode $embedded
                $sourceMatches = $false
                foreach ($sourceVariant in (Get-SourceCodeVariants -Path $sourceFile.FullName)) {
                    if ($embeddedNormalized -ieq $sourceVariant) { $sourceMatches = $true; break }
                }
                if (-not $sourceMatches) { [void]$mismatches.Add("code mismatch $($sourceFile.BaseName)") }
            }
            finally {
                Release-ComObject -Value $component
            }
        }
        if ($mismatches.Count -gt 0) { throw ('Source/book synchronization failed: ' + ($mismatches -join '; ')) }
    }
    finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject -Value $workbook }
        if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject -Value $excel }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Test-OpenXmlWorkbook {
    param([Parameter(Mandatory = $true)][string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $names = @($zip.Entries | ForEach-Object { $_.FullName })
        foreach ($required in @('[Content_Types].xml', 'xl/workbook.xml', 'xl/vbaProject.bin')) {
            if ($names -notcontains $required) { throw "Open XML part is missing: $required" }
        }
        $backslashPath = $names | Where-Object { $_ -match '\\' } | Select-Object -First 1
        if ($null -ne $backslashPath) { throw "Open XML entry uses backslash: $backslashPath" }
    }
    finally { $zip.Dispose() }
}

function Test-ExcelReadOnlyWorkbook {
    param([Parameter(Mandatory = $true)][string]$Path)
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AutomationSecurity = 1 } catch {}
        $workbook = $excel.Workbooks.Open($Path, 0, $true)
        if ($workbook.Worksheets.Count -lt 1) { throw 'Read-only Excel open produced no worksheets.' }
        if ($workbook.VBProject.VBComponents.Count -lt 1) { throw 'Read-only Excel open produced no VBA components.' }
    }
    finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject -Value $workbook }
        if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject -Value $excel }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    Wait-OfficeClosed
}

function Test-OfficeClosed {
    $processes = @(Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue)
    if ($processes.Count -gt 0) { throw ('Office is already running: ' + (($processes | ForEach-Object { $_.ProcessName }) -join ', ')) }
}

function Wait-OfficeClosed {
    param([int]$Attempts = 120)
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        if (-not (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 500
    }
    throw 'Office process remained running after a disposable gate.'
}

function Invoke-TestCase {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [string[]]$Arguments = @()
    )
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $logPath = Join-Path $logDirectory ($Id + '.log')
    $command = ($ScriptPath + ' ' + ($Arguments -join ' ')).Trim()
    try {
        $output = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ScriptPath @Arguments 2>&1 | ForEach-Object { $_.ToString() })
        $exitCode = if ($null -eq $LASTEXITCODE) { 0 } else { [int]$LASTEXITCODE }
        [IO.File]::WriteAllLines($logPath, $output, [Text.UTF8Encoding]::new($false))
        $message = if ($output.Count -gt 0) { $output[-1] } else { 'completed without output' }
        if ($exitCode -ne 0) { throw "exit_code=$exitCode; $message" }
        Add-GateResult -Id $Id -Status 'PASS' -ExitCode 0 -DurationMs $timer.ElapsedMilliseconds -Command $command -Message 'completed' -LogPath $logPath
        return $true
    }
    catch {
        $timer.Stop()
        if (-not (Test-Path -LiteralPath $logPath)) { [IO.File]::WriteAllText($logPath, $_.Exception.Message, [Text.UTF8Encoding]::new($false)) }
        Add-GateResult -Id $Id -Status 'FAIL' -ExitCode 30 -DurationMs $timer.ElapsedMilliseconds -Command $command -Message $_.Exception.Message -LogPath $logPath
        return $false
    }
    finally { $timer.Stop() }
}

function Write-GateReports {
    param([int]$Code)
    $report = [ordered]@{
        schema = 'createorder.release-gate.v1'
        generated_at = (Get-Date).ToString('o')
        mode = $Mode
        skip_manual = [bool]$SkipManual
        workbook = $WorkbookPath
        source_directory = $SourceDirectory
        output_directory = $OutputDirectory
        backup = $backupPath
        release = $releasePath
        exit_code = $Code
        first_failure = $firstFailure
        gates = @($gateResults)
    }
    $json = $report | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($reportJsonPath, $json, [Text.UTF8Encoding]::new($true))
    $lines = [Collections.Generic.List[string]]::new()
    [void]$lines.Add('# CreateOrder release-gate')
    [void]$lines.Add('')
    [void]$lines.Add("- Mode: $Mode")
    [void]$lines.Add("- Result: $(if ($Code -eq 0) { 'PASS' } else { 'FAIL' }) (exit $Code)")
    [void]$lines.Add("- Workbook: $WorkbookPath")
    [void]$lines.Add("- Backup: $(if ($backupPath) { $backupPath } else { 'not created in Verify mode' })")
    [void]$lines.Add("- Release: $(if ($releasePath) { $releasePath } else { 'not created' })")
    [void]$lines.Add('')
    [void]$lines.Add('| Gate | Status | Exit | Duration ms | Message |')
    [void]$lines.Add('|---|---:|---:|---:|---|')
    foreach ($item in $gateResults) { [void]$lines.Add("| $($item.id) | $($item.status) | $($item.exit_code) | $($item.duration_ms) | $($item.message -replace '\|','/') |") }
    [void]$lines.Add('')
    [void]$lines.Add('Manual owner checks are not inferred from automated tests.')
    [IO.File]::WriteAllLines($reportMarkdownPath, $lines, [Text.UTF8Encoding]::new($true))
}

try {
    if ($Mode -notin @('Verify', 'Release')) {
        Add-GateResult -Id 'parameters' -Status 'FAIL' -ExitCode 10 -Message "Invalid mode '$Mode'. Use Verify or Release."
        $overallCode = 10
        throw "Invalid mode '$Mode'."
    }
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath" }
    if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) { throw "Source directory not found: $SourceDirectory" }
    Test-OfficeClosed
    Add-GateResult -Id 'preflight-office' -Status 'PASS' -Message 'Excel and Word are closed.'

    $probeDirectory = Join-Path $reportDirectory 'probe'
    New-Item -ItemType Directory -Path $probeDirectory -Force | Out-Null
    $probeWorkbook = Join-Path $probeDirectory 'CreateOrder.gate.xlsm'
    Copy-Item -LiteralPath $WorkbookPath -Destination $probeWorkbook -Force
    Add-GateResult -Id 'input-and-probe' -Status 'PASS' -Message 'Disposable workbook copy created.'

    $timer = [Diagnostics.Stopwatch]::StartNew()
    try { Test-SourceBookSynchronization -BookPath $probeWorkbook; Add-GateResult -Id 'source-book-sync' -Status 'PASS' -DurationMs $timer.ElapsedMilliseconds -Message 'VBA source and embedded components match.' }
    catch { Add-GateResult -Id 'source-book-sync' -Status 'FAIL' -ExitCode 20 -DurationMs $timer.ElapsedMilliseconds -Message $_.Exception.Message; $overallCode = 20; throw }
    finally { $timer.Stop() }

    $timer = [Diagnostics.Stopwatch]::StartNew()
    try { Test-OpenXmlWorkbook -Path $probeWorkbook; Add-GateResult -Id 'openxml-probe' -Status 'PASS' -DurationMs $timer.ElapsedMilliseconds -Message 'Workbook ZIP contains required parts and normalized paths.' }
    catch { Add-GateResult -Id 'openxml-probe' -Status 'FAIL' -ExitCode 20 -DurationMs $timer.ElapsedMilliseconds -Message $_.Exception.Message; $overallCode = 20; throw }
    finally { $timer.Stop() }

    $testCases = @(
        @{ Id = 'personnel-v2-designer'; Script = Join-Path $projectRoot 'tools\Test-PersonnelActionWizardV2Designer.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory, '-ExpectedActiveVersion', 'V2') },
        @{ Id = 'personnel-preview'; Script = Join-Path $projectRoot 'tools\Test-PersonnelActionPreviewSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory) },
        @{ Id = 'personnel-v2-e2e'; Script = Join-Path $projectRoot 'tools\Test-PersonnelActionWizardV2Safe.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory) },
        @{ Id = 'integrity-fixture'; Script = Join-Path $projectRoot 'tools\Test-PersonnelDataIntegritySafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory) },
        @{ Id = 'integrity-designer'; Script = Join-Path $projectRoot 'tools\Test-DataIntegrityCenterDesigner.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory) },
        @{ Id = 'history-center'; Script = Join-Path $projectRoot 'tools\Test-PersonnelHistoryCenterSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory) },
        @{ Id = 'history-center-designer'; Script = Join-Path $projectRoot 'tools\Test-PersonnelHistoryCenterDesigner.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory) },
        @{ Id = 'grouped-personnel-order'; Script = Join-Path $projectRoot 'tools\Test-GroupedPersonnelOrderSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-SourceDirectory', $SourceDirectory) },
        @{ Id = 'personnel-ribbon'; Script = Join-Path $projectRoot 'tools\Test-PersonnelRibbonSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook) },
        @{ Id = 'personnel-action'; Script = Join-Path $projectRoot 'tools\Test-PersonnelActionWizardSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook) },
        @{ Id = 'personnel-events'; Script = Join-Path $projectRoot 'Test-PersonnelEvents.ps1'; Args = @('-WorkbookPath', $probeWorkbook) },
        @{ Id = 'enrollment-compact-ui'; Script = Join-Path $projectRoot 'tools\Test-EnrollmentCompactUiSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook) },
        @{ Id = 'enrollment-fizo'; Script = Join-Path $projectRoot 'tools\Test-EnrollmentFizoReferenceSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook) },
        @{ Id = 'enrollment-tariff'; Script = Join-Path $projectRoot 'tools\Test-EnrollmentTariffReferenceSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook) },
        @{ Id = 'enrollment-medal'; Script = Join-Path $projectRoot 'tools\Test-EnrollmentMedalReferenceSafe.ps1'; Args = @('-WorkbookPath', $probeWorkbook) },
        @{ Id = 'fio-declension'; Script = Join-Path $projectRoot 'Test-FIODeclension.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-HelperPath', (Join-Path $SourceDirectory 'mdlHelper.bas')) },
        @{ Id = 'zp12-validation'; Script = Join-Path $projectRoot 'Test-ZP12Validation.ps1'; Args = @('-WorkbookPath', $probeWorkbook, '-ModulePath', (Join-Path $SourceDirectory 'mdlZP12Validation.bas'), '-HelperPath', (Join-Path $SourceDirectory 'mdlHelper.bas'), '-RibbonHandlersPath', (Join-Path $SourceDirectory 'mdlRibbonHandlers.bas')) },
        @{ Id = 'full-acceptance'; Script = Join-Path $projectRoot 'Test-PaymentsEnrollmentAcceptance.ps1'; Args = @() }
    )
    Wait-OfficeClosed
    foreach ($test in $testCases) {
        if (-not (Test-Path -LiteralPath $test.Script -PathType Leaf)) {
            Add-GateResult -Id $test.Id -Status 'FAIL' -ExitCode 30 -Message "Test script missing: $($test.Script)"; $overallCode = 30; throw "Test script missing: $($test.Script)"
        }
        if (-not (Invoke-TestCase -Id $test.Id -ScriptPath $test.Script -Arguments $test.Args)) { $overallCode = 30; throw "Gate test failed: $($test.Id)" }
        Wait-OfficeClosed
    }

    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        if ($Mode -eq 'Release') {
            $backupDirectory = Join-Path $projectRoot "CreateOrderBackups\release-gate-$stamp"
            New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
            $backupPath = Join-Path $backupDirectory ([IO.Path]::GetFileName($WorkbookPath) + '.before-release-gate.xlsm')
            Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
            $buildScript = Join-Path $projectRoot 'Build-Release.ps1'
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $buildScript -SourceFile $probeWorkbook -OutputDirectory $OutputDirectory -SkipGate 2>&1 | Out-File -LiteralPath (Join-Path $logDirectory 'build-release.log') -Encoding utf8
            if ($LASTEXITCODE -ne 0) { throw "Build-Release failed with exit code $LASTEXITCODE." }
            $releasePath = Get-ChildItem -LiteralPath $OutputDirectory -Filter 'CreateOrder_Release_*.xlsm' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
            if ([string]::IsNullOrWhiteSpace($releasePath)) { throw 'Build-Release did not create a release artifact.' }
            Test-OpenXmlWorkbook -Path $releasePath
            Test-ExcelReadOnlyWorkbook -Path $releasePath
            Add-GateResult -Id 'release-artifact' -Status 'PASS' -DurationMs $timer.ElapsedMilliseconds -Message 'Release artifact passed Open XML and Excel read-only verification.' -LogPath (Join-Path $logDirectory 'build-release.log')
        } else {
            Add-GateResult -Id 'release-artifact' -Status 'PASS' -DurationMs $timer.ElapsedMilliseconds -Message 'Verify mode does not create a release artifact.'
        }
    }
    catch { Add-GateResult -Id 'release-artifact' -Status 'FAIL' -ExitCode 40 -DurationMs $timer.ElapsedMilliseconds -Message $_.Exception.Message -LogPath (Join-Path $logDirectory 'build-release.log'); $overallCode = 40; throw }
    finally { $timer.Stop() }

    if ($SkipManual) { Add-GateResult -Id 'manual-owner-gate' -Status 'WARN' -ExitCode 0 -Message 'Manual visual/owner checks explicitly skipped.' }
    else { Add-GateResult -Id 'manual-owner-gate' -Status 'MANUAL_REQUIRED' -ExitCode 50 -Message 'Visual form, Word layout and owner decisions still require manual acceptance.'; $overallCode = 50 }
}
catch {
    if ($overallCode -eq 0) {
        $overallCode = 10
        Add-GateResult -Id 'unhandled-error' -Status 'FAIL' -ExitCode $overallCode -Message $_.Exception.Message
    }
}
finally {
    Write-GateReports -Code $overallCode
    Write-Output "CREATEORDER_RELEASE_GATE|mode=$Mode|exit=$overallCode|report=$reportMarkdownPath|json=$reportJsonPath|backup=$backupPath|release=$releasePath"
}

exit $overallCode
