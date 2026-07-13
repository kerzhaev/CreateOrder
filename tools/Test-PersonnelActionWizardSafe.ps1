param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_personnel_action_wizard_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_personnel_action_wizard_test.xlsm"
$moduleDirectory = Join-Path $workspace "CreateOrder.xlsm.modules"

function Read-VbaText([string]$Path) {
    [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = Read-VbaText $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($code)
}

function Import-UserForm($Workbook, [string]$FormName, [string]$FormPath) {
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item($FormName)) } catch {}
    $component = $Workbook.VBProject.VBComponents.Import($FormPath)
    if ($component.Name -ne $FormName) { throw "Imported form name mismatch: $($component.Name)" }
}

New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
Copy-Item -LiteralPath $WorkbookPath -Destination $testWorkbookPath -Force
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 1 } catch {}
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)
    Import-CodeModuleText $workbook "ModuleLocalization" (Join-Path $moduleDirectory "ModuleLocalization.bas")
    Import-CodeModuleText $workbook "mdlPersonnelEvents" (Join-Path $moduleDirectory "mdlPersonnelEvents.bas")
    Import-UserForm $workbook "frmPersonnelActionWizard" (Join-Path $moduleDirectory "frmPersonnelActionWizard.frm")

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("personnel_action_wizard_probe")) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = "personnel_action_wizard_probe"
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function ProbePersonnelActionWizard() As String
    Dim enrollmentID As String, transferID As String, outputPath As String, employeeID As String, currentState As Object
    On Error GoTo Failed
    mdlPersonnelEvents.ResetPersonnelEventInput
    mdlPersonnelEvents.SetPersonnelWizardValue "event_type", "ENROLLMENT"
    mdlPersonnelEvents.SetPersonnelWizardValue "event_date", DateSerial(2026, 7, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "effective_date", DateSerial(2026, 7, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "order_reference", "WIZ-ENROLL-001"
    mdlPersonnelEvents.SetPersonnelWizardValue "basis_text", "Wizard test enrollment"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_fio", "Wizard Test Employee"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_personal_number", "WIZ-001"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_rank", "Private"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_position", "Initial position"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_section", "Initial section"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_military_unit", "Test unit"
    enrollmentID = mdlPersonnelEvents.SavePersonnelEventInput(False)
    employeeID = CStr(mdlPersonnelEvents.GetPersonnelWizardValue("employee_id"))
    If employeeID = "" Then Err.Raise 699, , "Enrollment did not create EmployeeID"
    mdlPersonnelEvents.PrepareNewPersonnelAction "TRANSFER"
    Load frmPersonnelActionWizard
    If frmPersonnelActionWizard.Controls("txt_employee_id") Is Nothing Then Err.Raise 700, , "Employee field missing"
    If frmPersonnelActionWizard.Controls("txt_destination_location") Is Nothing Then Err.Raise 701, , "Destination field missing"
    If frmPersonnelActionWizard.Controls("txt_status") Is Nothing Then Err.Raise 702, , "Status field missing"
    If frmPersonnelActionWizard.btnImportResponse.Caption = "" Then Err.Raise 703, , "Save button caption missing"
    frmPersonnelActionWizard.Controls("txt_employee_id").Value = employeeID
    mdlPersonnelEvents.SetPersonnelWizardValue "employee_id", employeeID
    If Not mdlPersonnelEvents.LoadPersonnelWizardCurrentState() Then Err.Raise 704, , "Wizard could not load current state"
    frmPersonnelActionWizard.Controls("txt_event_date").Value = "02.07.2026"
    frmPersonnelActionWizard.Controls("txt_effective_date").Value = "02.07.2026"
    frmPersonnelActionWizard.Controls("txt_order_reference").Value = "WIZ-TRANSFER-001"
    frmPersonnelActionWizard.Controls("txt_basis_text").Value = "Wizard test transfer"
    frmPersonnelActionWizard.Controls("txt_new_position").Value = "Transferred position"
    transferID = frmPersonnelActionWizard.SaveAction()
    If transferID = "" Then Err.Raise 705, , "Wizard save did not return EventID"
    Set currentState = mdlPersonnelEvents.GetCurrentPersonnelState(employeeID)
    If CStr(currentState("position")) <> "Transferred position" Then Err.Raise 706, , "Wizard save did not update current state"
    outputPath = frmPersonnelActionWizard.ExportAction()
    If outputPath = "" Then Err.Raise 707, , "Wizard export did not return output path"
    Unload frmPersonnelActionWizard
    ProbePersonnelActionWizard = "OK"
    Exit Function
Failed:
    ProbePersonnelActionWizard = "FAILED: " & Err.Description
End Function
"@)
    $result = $excel.Run("'$($workbook.Name)'!ProbePersonnelActionWizard")
    if ($result -ne "OK") { throw $result }
    $workbook.Close($false); $workbook = $null
    $excel.Quit(); $excel = $null
    Write-Output "Personnel action wizard safe acceptance passed."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
