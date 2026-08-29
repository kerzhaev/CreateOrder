[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [switch]$SkipPreflight
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-InstallLog {
    param([string]$Level, [string]$Message, [hashtable]$Context = @{})
    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = 'Install-PersonnelOrderAndLegacyBikFixSafe'
        message = $Message
    }
    foreach ($key in $Context.Keys) { $payload[$key] = $Context[$key] }
    Write-Host ($payload | ConvertTo-Json -Compress -Depth 4)
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

if (-not ('PersonnelFixInstallNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PersonnelFixInstallNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId([object]$ExcelApplication) {
    [uint32]$processId = 0
    [void][PersonnelFixInstallNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
    return [int]$processId
}

function Stop-OwnedExcelProcessIfNeeded([int]$ProcessId) {
    if ($ProcessId -le 0) { return }
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 250
    }
    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -eq 'EXCEL') {
        Write-InstallLog WARN 'Excel did not exit after COM Quit; stopping only the installer-owned process.' @{ processId = $ProcessId }
        Stop-Process -Id $ProcessId -Force
    }
}

function Read-VbaText([string]$Path) {
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText([object]$Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = Read-VbaText $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $null
    $codeModule = $null
    try {
        $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
        $codeModule = $component.CodeModule
        if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
        $codeModule.AddFromString($code)
    }
    finally {
        Release-ComObject $codeModule
        Release-ComObject $component
    }
}

function Ensure-LocalizationEntry([object]$Workbook, [string]$Key, [string]$RussianText) {
    $sheet = $null
    try {
        $sheet = $Workbook.Worksheets.Item('Localization')
        $lastRow = [int]$sheet.Cells($sheet.Rows.Count, 1).End(-4162).Row
        for ($row = 2; $row -le $lastRow; $row++) {
            if ([string]::Equals(([string]$sheet.Cells($row, 1).Value2).Trim(), $Key, [StringComparison]::OrdinalIgnoreCase)) {
                if ([string]::IsNullOrWhiteSpace([string]$sheet.Cells($row, 2).Value2)) {
                    $sheet.Cells($row, 2).Value2 = $RussianText
                    Write-InstallLog INFO 'Filled an empty Russian localization value.' @{ key = $Key; row = $row }
                }
                else {
                    Write-InstallLog DEBUG 'Localization entry already exists; preserving workbook value.' @{ key = $Key; row = $row }
                }
                return
            }
        }
        $newRow = [Math]::Max(2, $lastRow + 1)
        $sheet.Cells($newRow, 1).Value2 = $Key
        $sheet.Cells($newRow, 2).Value2 = $RussianText
        Write-InstallLog INFO 'Added the required localization entry.' @{ key = $Key; row = $newRow }
    }
    finally {
        Release-ComObject $sheet
    }
}

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
if (-not $resolvedWorkbook.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook must be inside the CreateOrder project: $resolvedWorkbook"
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel is open. Close all Excel windows before installing the personnel-order and legacy-BIK fix.'
}

if (-not $SkipPreflight) {
    Write-InstallLog INFO 'Running isolated preflight before changing the working workbook.' @{ workbook = $resolvedWorkbook }
    & (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardSafe.ps1') -WorkbookPath $resolvedWorkbook
    if (-not $?) { throw 'Personnel action preflight failed.' }
    for ($attempt = 1; $attempt -le 40 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) {
        Start-Sleep -Milliseconds 250
    }
    if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
        throw 'Office remained open after personnel action preflight; the working workbook was not changed.'
    }
    & (Join-Path $projectRoot 'Test-PaymentsEnrollmentAcceptance.ps1') -WorkbookPath $resolvedWorkbook
    if (-not $?) { throw 'Enrollment acceptance preflight failed.' }
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\personnel-order-legacy-bik-fixed-$stamp")
$backupPath = Join-Path $backupDirectory 'CreateOrder.before-personnel-order-legacy-bik-fix.xlsm'
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupPath
Write-InstallLog INFO 'Created the pre-installation workbook backup.' @{ backup = $backupPath }

$moduleDirectory = Join-Path $projectRoot 'CreateOrder.xlsm.modules'
$excel = $null
$excelProcessId = 0
$book = $null
$components = $null
$oldForm = $null
$importedForm = $null
$probe = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)
    if ($book.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot install the fix.' }

    Import-CodeModuleText $book 'ModuleLocalization' (Join-Path $moduleDirectory 'ModuleLocalization.bas')
    Import-CodeModuleText $book 'mdlEnrollmentWorkflow' (Join-Path $moduleDirectory 'mdlEnrollmentWorkflow.bas')
    Import-CodeModuleText $book 'mdlPersonnelEventOrderExport' (Join-Path $moduleDirectory 'mdlPersonnelEventOrderExport.bas')
    Ensure-LocalizationEntry $book 'personnel.wizard.material_assistance_status' 'Материальная помощь за год'
    Ensure-LocalizationEntry $book 'personnel.wizard.main_leave_status' 'Основной отпуск за год'
    Ensure-LocalizationEntry $book 'personnel.wizard.additional_leave_status' 'Дополнительный отпуск за год'

    $components = $book.VBProject.VBComponents
    $oldForm = $components.Item('frmPersonnelActionWizard')
    $components.Remove($oldForm)
    Release-ComObject $oldForm
    $oldForm = $null
    $importedForm = $components.Import((Join-Path $moduleDirectory 'frmPersonnelActionWizard.frm'))
    if ($importedForm.Name -ne 'frmPersonnelActionWizard' -or $importedForm.Type -ne 3) {
        throw "Personnel form import verification failed: name=$($importedForm.Name), type=$($importedForm.Type)"
    }

    $probe = $components.Add(1)
    $probe.Name = 'modPersonnelFixInstallProbe'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Sub VerifyPersonnelFixInstall()
    ModuleLocalization.ResetLocalizationCache
    mdlPersonnelEvents.ResetPersonnelEventInput
    mdlPersonnelEvents.SetPersonnelWizardValue "event_type", "EXCLUSION"
    Load frmPersonnelActionWizard
    If frmPersonnelActionWizard.Controls("txt_material_assistance_status") Is Nothing Then Err.Raise 5, , "material assistance field missing"
    If frmPersonnelActionWizard.Controls("txt_main_leave_status") Is Nothing Then Err.Raise 5, , "main leave field missing"
    If frmPersonnelActionWizard.Controls("txt_additional_leave_status") Is Nothing Then Err.Raise 5, , "additional leave field missing"
    Unload frmPersonnelActionWizard
End Sub
'@)
    $excel.Run("'$($book.Name)'!modPersonnelFixInstallProbe.VerifyPersonnelFixInstall")
    $components.Remove($probe)
    Release-ComObject $probe
    $probe = $null

    $book.Save()
    Write-InstallLog INFO 'Installed the checked modules and personnel form.' @{ workbook = $resolvedWorkbook; backup = $backupPath }
}
catch {
    Write-InstallLog ERROR 'Installation failed; the pre-installation backup is available.' @{ error = $_.Exception.Message; backup = $backupPath }
    throw
}
finally {
    if ($probe -and $components) { try { $components.Remove($probe) } catch {} }
    if ($book) { try { $book.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $probe
    Release-ComObject $importedForm
    Release-ComObject $components
    Release-ComObject $book
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded $excelProcessId
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel remained running after installation.' }
Write-InstallLog INFO 'Running post-install personnel and localization verification.' @{ workbook = $resolvedWorkbook }
& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardSafe.ps1') -WorkbookPath $resolvedWorkbook -RequireInstalledLocalization
if (-not $?) { throw 'Post-install personnel and localization verification failed.' }
for ($attempt = 1; $attempt -le 40 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
    throw 'Office remained open after post-install verification.'
}
Write-InstallLog INFO 'Personnel-order and legacy-BIK installation completed.' @{ workbook = $resolvedWorkbook; backup = $backupPath }
