[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmPersonnelActionWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-ActivationLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$Context = @{}
    )
    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = 'Activate-PersonnelActionWizardV2Safe'
        message = $Message
    }
    foreach ($key in $Context.Keys) { $payload[$key] = $Context[$key] }
    $line = $payload | ConvertTo-Json -Compress -Depth 5
    if ($Level -eq 'DEBUG') { Write-Verbose $line }
    elseif ($Level -eq 'WARN') { Write-Warning $line }
    elseif ($Level -eq 'ERROR') { Write-Error $line -ErrorAction Continue }
    else { Write-Host $line }
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

if (-not ('PersonnelV2ActivationNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PersonnelV2ActivationNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][PersonnelV2ActivationNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
    return [int]$processId
}

function Stop-OwnedExcelProcessIfNeeded {
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return }
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 250
    }
    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -eq 'EXCEL') {
        Write-ActivationLog WARN 'Excel did not exit after COM Quit; stopping only the activation-owned process.' @{ processId = $ProcessId }
        Stop-Process -Id $ProcessId -Force
    }
}

function Wait-ForOfficeShutdown {
    param([int]$TimeoutSeconds = 15)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $office = @(Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue)
        if ($office.Count -eq 0) { return }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    $remaining = @(Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue | ForEach-Object { "$($_.ProcessName) (PID $($_.Id))" })
    throw ('Office did not exit after a safe wait: ' + ($remaining -join ', '))
}

function Read-VbaText {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Set-WorkbookCodeModule {
    param(
        [Parameter(Mandatory = $true)][string]$TargetWorkbook,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $true)][string]$ModulePath,
        [Parameter(Mandatory = $true)][string]$V2FormName
    )

    $excel = $null
    $excelProcessId = 0
    $book = $null
    $component = $null
    $codeModule = $null
    $v1 = $null
    $v2 = $null
    try {
        Write-ActivationLog INFO 'Opening workbook to switch personnel action routing.' @{ workbook = $TargetWorkbook; module = $ModuleName }
        $excel = New-Object -ComObject Excel.Application
        $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1
        $book = $excel.Workbooks.Open($TargetWorkbook, 0, $false)
        if ($book.ReadOnly) { throw "Workbook opened read-only: $TargetWorkbook" }
        if ($book.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot activate personnel V2.' }

        $v1 = $book.VBProject.VBComponents.Item('frmPersonnelActionWizard')
        $v2 = $book.VBProject.VBComponents.Item($V2FormName)
        if ($v1.Type -ne 3 -or $v2.Type -ne 3) { throw 'V1 and V2 must both remain UserForm components during activation.' }

        $code = Read-VbaText -Path $ModulePath
        $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
        $component = $book.VBProject.VBComponents.Item($ModuleName)
        if ($component.Type -ne 1) { throw "Unexpected component type for $ModuleName`: $($component.Type)" }
        $codeModule = $component.CodeModule
        if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
        $codeModule.AddFromString($code)

        $installedCode = $codeModule.Lines(1, $codeModule.CountOfLines)
        if (-not $installedCode.Contains(($V2FormName + '.Show'))) { throw 'V2 route is missing after module replacement.' }
        if ($installedCode.Contains('frmPersonnelActionWizard.Show')) { throw 'V1 route is still active after module replacement.' }

        $book.Save()
        Write-ActivationLog INFO 'Personnel action routing switched and workbook saved.' @{ workbook = $TargetWorkbook; activeForm = $V2FormName; rollbackForm = 'frmPersonnelActionWizard' }
    } finally {
        if ($book) { $book.Close($false) }
        if ($excel) { $excel.Quit() }
        Release-ComObject $codeModule
        Release-ComObject $component
        Release-ComObject $v2
        Release-ComObject $v1
        Release-ComObject $book
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
    }
}

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
if (-not $resolvedWorkbook.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook must be inside the CreateOrder project: $resolvedWorkbook"
}
if (-not $resolvedSource.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Source directory must be inside the CreateOrder project: $resolvedSource"
}
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
    throw 'Excel or Word is running. Save and close Office before activating personnel V2.'
}

$modulePath = Join-Path $resolvedSource 'mdlPersonnelEvents.bas'
if (-not (Test-Path -LiteralPath $modulePath)) { throw "Missing personnel events source: $modulePath" }
$sourceCode = Read-VbaText -Path $modulePath
$v2RouteCount = [regex]::Matches($sourceCode, [regex]::Escape($TargetComponentName + '.Show')).Count
$v1RouteCount = [regex]::Matches($sourceCode, '(?<!V2)frmPersonnelActionWizard\.Show').Count
if ($v2RouteCount -ne 2 -or $v1RouteCount -ne 0) {
    throw "Unexpected source routing: V2 routes=$v2RouteCount, V1 routes=$v1RouteCount. Expected 2 and 0."
}
Write-ActivationLog DEBUG 'Validated source routing before activation.' @{ v2Routes = $v2RouteCount; v1Routes = $v1RouteCount }

& (Join-Path $PSScriptRoot 'Install-PersonnelActionWizardV2Safe.ps1') `
    -WorkbookPath $resolvedWorkbook `
    -SourceDirectory $resolvedSource `
    -TargetComponentName $TargetComponentName

& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardV2Safe.ps1') `
    -WorkbookPath $resolvedWorkbook `
    -SourceDirectory $resolvedSource `
    -TargetComponentName $TargetComponentName
& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardSafe.ps1') -WorkbookPath $resolvedWorkbook
Wait-ForOfficeShutdown

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stagingDirectory = Join-Path $projectRoot ("Trash\personnel-action-v2-activation-probe-$stamp")
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\personnel-action-v2-activated-$stamp")
New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$stagingWorkbook = Join-Path $stagingDirectory 'CreateOrder.personnel-action-v2-activation-probe.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-personnel-action-v2-activation.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $stagingWorkbook
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook
Write-ActivationLog INFO 'Prepared activation staging workbook and safety backup.' @{ stagingWorkbook = $stagingWorkbook; backupWorkbook = $backupWorkbook }

Set-WorkbookCodeModule -TargetWorkbook $stagingWorkbook -ModuleName 'mdlPersonnelEvents' -ModulePath $modulePath -V2FormName $TargetComponentName
Wait-ForOfficeShutdown
& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardV2Designer.ps1') `
    -WorkbookPath $stagingWorkbook `
    -SourceDirectory $resolvedSource `
    -TargetComponentName $TargetComponentName `
    -ExpectedActiveVersion V2

Set-WorkbookCodeModule -TargetWorkbook $resolvedWorkbook -ModuleName 'mdlPersonnelEvents' -ModulePath $modulePath -V2FormName $TargetComponentName
Wait-ForOfficeShutdown
try {
    & (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardV2Designer.ps1') `
        -WorkbookPath $resolvedWorkbook `
        -SourceDirectory $resolvedSource `
        -TargetComponentName $TargetComponentName `
        -ExpectedActiveVersion V2
} catch {
    Write-ActivationLog ERROR 'Post-activation verification failed; restoring the safety backup.' @{ workbook = $resolvedWorkbook; backupWorkbook = $backupWorkbook; error = $_.Exception.Message }
    Copy-Item -LiteralPath $backupWorkbook -Destination $resolvedWorkbook -Force
    throw
}

Write-ActivationLog INFO 'Personnel action V2 activation completed.' @{ workbook = $resolvedWorkbook; activeForm = $TargetComponentName; rollbackForm = 'frmPersonnelActionWizard'; backup = $backupWorkbook }
