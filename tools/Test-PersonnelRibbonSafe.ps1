param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_personnel_ribbon_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_personnel_ribbon_test.xlsm"

function Read-VbaText([string]$Path) {
    [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
    $codeModule = $component.CodeModule
    if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
    $codeModule.AddFromString($code)
}

New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
Copy-Item -LiteralPath $WorkbookPath -Destination $testWorkbookPath -Force

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # Enable macros only inside this disposable workbook copy so its VBA can be exercised.
    try { $excel.AutomationSecurity = 1 } catch {}
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)

    Import-CodeModuleText $workbook "ModuleLocalization" (Join-Path $workspace "CreateOrder.xlsm.modules\ModuleLocalization.bas")
    Import-CodeModuleText $workbook "mdlRibbonHandlers" (Join-Path $workspace "CreateOrder.xlsm.modules\mdlRibbonHandlers.bas")

    $result = $excel.Run("'$($workbook.Name)'!mdlRibbonHandlers.GetRibbonUiTextById", "personnelActionsGroup", "label")
    $expected = -join @(1050,1072,1076,1088,1086,1074,1099,1077,32,1076,1077,1081,1089,1090,1074,1080,1103 | ForEach-Object { [char]$_ })
    if ($result -ne $expected) { throw "Personnel ribbon localization failed. Actual: $result" }

    $result = $excel.Run("'$($workbook.Name)'!mdlRibbonHandlers.GetRibbonUiTextById", "openPersonnelActionsMenu", "label")
    $expected = -join @(1054,1090,1082,1088,1099,1090,1100,32,1082,1072,1076,1088,1086,1074,1099,1077,32,1076,1077,1081,1089,1090,1074,1080,1103 | ForEach-Object { [char]$_ })
    if ($result -ne $expected) { throw "Personnel action menu localization failed. Actual: $result" }

    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    Write-Output "Personnel ribbon safe acceptance passed."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
