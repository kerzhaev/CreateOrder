param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot

function Read-VbaText([string]$Path) {
    [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = Read-VbaText $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $module = $Workbook.VBProject.VBComponents.Item($ModuleName).CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($code)
}

$openExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($openExcel.Count -gt 0) {
    throw "Excel is open. Close all Excel windows before importing VBA into the working workbook."
}
if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupDirectory = Join-Path $workspace ("CreateOrderBackups\enrollment-compact-ui-installed-" + $stamp)
$backupPath = Join-Path $backupDirectory "CreateOrder.before-enrollment-compact-ui.xlsm"
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 3 } catch {}
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)

    Import-CodeModuleText $workbook "mdlEnrollmentWorkflow" (Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentWorkflow.bas")
    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("frmEnrollmentWizard")) } catch {}
    $form = $workbook.VBProject.VBComponents.Import((Join-Path $workspace "CreateOrder.xlsm.modules\frmEnrollmentWizard.frm"))
    if ($form.Type -ne 3) { throw "Enrollment form was imported as component type $($form.Type), expected 3." }

    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    Write-Output "Enrollment compact UI installed. Backup: $backupPath"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
