param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_enrollment_medal_reference_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_enrollment_medal_reference_test.xlsm"

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$Path) {
    $code = [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $module = $Workbook.VBProject.VBComponents.Item($ModuleName).CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($code)
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
    Import-CodeModuleText $workbook "mdlEnrollmentWorkflow" (Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentWorkflow.bas")
    Import-CodeModuleText $workbook "mdlPersonnelEvents" (Join-Path $workspace "CreateOrder.xlsm.modules\mdlPersonnelEvents.bas")
    $formPath = Join-Path $workspace "CreateOrder.xlsm.modules\frmEnrollmentWizard.frm"
    $components = $workbook.VBProject.VBComponents
    try { $components.Remove($components.Item("frmEnrollmentWizard")) } catch {}
    $components.Import($formPath) | Out-Null

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("enr_medal_probe")) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = "enr_medal_probe"
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function ProbeEnrollmentMedalReferences() As String
    Dim values As Collection
    Dim record As Object
    Dim stateData As Object
    Dim eventID As String
    Dim employeeID As String
    Dim medalCode As String
    Dim rowNum As Long
    Dim i As Long
    Dim totalAmount As Long
    Load frmEnrollmentWizard
    Unload frmEnrollmentWizard
    mdlEnrollmentWorkflow.EnsureEnrollmentReferenceData
    Set values = mdlEnrollmentWorkflow.GetEnrollmentReferenceValues("ACHIEVEMENT")
    If values.Count <> 4 Then
        ProbeEnrollmentMedalReferences = "FAILED: expected four medal references"
        Exit Function
    End If
    For i = 1 To values.Count
        totalAmount = totalAmount + CLng(mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("ACHIEVEMENT", CStr(values(i))))
        If mdlEnrollmentWorkflow.GetEnrollmentReferenceCode("ACHIEVEMENT", CStr(values(i))) = "" Then
            ProbeEnrollmentMedalReferences = "FAILED: medal reference code is missing"
            Exit Function
        End If
    Next i
    If totalAmount <> 80 Then
        ProbeEnrollmentMedalReferences = "FAILED: medal percentages must be 30, 20, 20 and 10"
        Exit Function
    End If
    Set record = CreateObject("Scripting.Dictionary")
    record("enrollment_id") = "MEDAL-REFERENCE-PROBE"
    record("fio") = "Medal reference probe"
    record("personal_number") = "MRP-001"
    record("rank") = ""
    record("position") = ""
    record("section") = ""
    record("military_unit") = ""
    record("vus") = ""
    record("tariff_rank") = ""
    record("position_salary") = ""
    record("rank_salary") = ""
    record("service_category") = "CONTRACT"
    record("contract_kind") = ""
    record("contract_basis") = ""
    record("order_date") = DateSerial(2026, 1, 15)
    record("enroll_date") = DateSerial(2026, 1, 15)
    record("duty_start_date") = DateSerial(2026, 1, 15)
    record("order_number") = "MEDAL-PROBE"
    record("basis_section1") = "test"
    record("achievement_param") = CStr(values(1))
    record("achievement_award_date") = DateSerial(2026, 1, 15)
    record("achievement_document_reference") = "AWARD-15"
    medalCode = mdlEnrollmentWorkflow.GetEnrollmentReferenceCode("ACHIEVEMENT", CStr(values(1)))
    eventID = mdlPersonnelEvents.EnsureEnrollmentPersonnelEvent(record)
    If eventID = "" Then
        ProbeEnrollmentMedalReferences = "FAILED: enrollment event was not created"
        Exit Function
    End If
    For rowNum = 2 To ThisWorkbook.Worksheets("PersonnelEvents").Cells(ThisWorkbook.Worksheets("PersonnelEvents").Rows.Count, 1).End(xlUp).Row
        If CStr(ThisWorkbook.Worksheets("PersonnelEvents").Cells(rowNum, 1).Value) = eventID Then
            employeeID = CStr(ThisWorkbook.Worksheets("PersonnelEvents").Cells(rowNum, 2).Value)
            Exit For
        End If
    Next rowNum
    Set stateData = mdlPersonnelEvents.GetCurrentPersonnelState(employeeID)
    If stateData.Count = 0 Then
        ProbeEnrollmentMedalReferences = "FAILED: current state is missing"
        Exit Function
    End If
    If CStr(stateData("medal_code")) <> medalCode Or CStr(stateData("medal_award_document_reference")) <> "AWARD-15" Then
        ProbeEnrollmentMedalReferences = "FAILED: medal data was not transferred to current state"
        Exit Function
    End If
    ProbeEnrollmentMedalReferences = "OK"
End Function
"@)
    $result = $excel.Run("'$($workbook.Name)'!ProbeEnrollmentMedalReferences")
    if ($result -ne "OK") { throw $result }
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    Write-Output "Enrollment medal reference safe acceptance passed."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
