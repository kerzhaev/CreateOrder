[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmPersonnelHistoryCenter'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-Code([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    try { return ([Text.UTF8Encoding]::new($false, $true)).GetString($bytes) } catch { return [Text.Encoding]::GetEncoding(1251).GetString($bytes) }
}

function Import-Code([object]$Workbook, [string]$Name, [string]$Path) {
    $code = Read-Code $Path
    $code = [regex]::Replace($code, '(?m)^Attribute .*\r?\n', '')
    $component = $null
    try { $component = $Workbook.VBProject.VBComponents.Item($Name) } catch {}
    if ($null -eq $component) { $component = $Workbook.VBProject.VBComponents.Add(1); $component.Name = $Name }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    [void]$module.AddFromString($code)
}

function Release-Com([object]$Value) { if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {} } }

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word is running. Close Office before the history center designer test.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$formPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$frxPath = Join-Path $resolvedSource ($TargetComponentName + '.frx')
$manifestPath = Join-Path $resolvedSource ($TargetComponentName + '.layout.csv')
foreach ($path in @($formPath, $frxPath, $manifestPath)) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "History center designer artifact is missing: $path" } }
$formText = Read-Code $formPath
if ($formText -match 'Controls\.Add|\.Left\s*=|\.Top\s*=|\.Width\s*=|\.Height\s*=') { throw 'History center form contains runtime control creation or geometry.' }
foreach ($handler in @('btnSearch_Click', 'btnOpenDocument_Click', 'btnRepeatExport_Click', 'btnPrepareCorrection_Click')) { if ($formText -notmatch ('Private Sub ' + $handler + '\(\)')) { throw "History center form is missing $handler." } }
$manifest = @(Import-Csv -LiteralPath $manifestPath)
if ($manifest.Count -ne 16) { throw "Expected 16 history center controls, got $($manifest.Count)." }
if (@($manifest.Name | Sort-Object -Unique).Count -ne 16) { throw 'History center control names are not unique.' }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testDirectory = Join-Path $projectRoot "Trash\personnel-history-center-designer-test-$stamp"
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$testWorkbook = Join-Path $testDirectory 'CreateOrder.history-center-designer.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $testWorkbook -Force

$excel = $null
$workbook = $null
$formComponent = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($testWorkbook, 0, $false)
    Import-Code $workbook 'mdlPersonnelHistoryCenter' (Join-Path $resolvedSource 'mdlPersonnelHistoryCenter.bas')
    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item($TargetComponentName)) } catch {}
    $formComponent = $workbook.VBProject.VBComponents.Import($formPath)
    if ($formComponent.Name -ne $TargetComponentName) { throw "Imported form has unexpected name $($formComponent.Name)." }
    $designer = $formComponent.Designer
    foreach ($row in $manifest) {
        $control = $designer.Controls.Item($row.Name)
        if ($null -eq $control) { throw "Designer control is missing: $($row.Name)." }
    }
    $timeline = $designer.Controls.Item('txtTimeline')
    if (-not [bool]$timeline.Locked) { throw 'Timeline textbox is not locked.' }
    if (-not [bool]$timeline.MultiLine) { throw 'Timeline textbox is not multiline.' }
    if ([int]$timeline.ScrollBars -ne 3) { throw 'Timeline textbox has no both scrollbars.' }
    $workbook.Close($false); Release-Com $workbook; $workbook = $null
    $excel.Quit(); Release-Com $excel; $excel = $null
    Write-Output "PERSONNEL_HISTORY_CENTER_DESIGNER_TEST_OK|controls=$($manifest.Count)|form=$formPath|manifest=$manifestPath"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-Com $workbook }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
for ($attempt = 1; $attempt -le 120 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after the history center designer test.' }
