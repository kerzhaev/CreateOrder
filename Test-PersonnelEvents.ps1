param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "CreateOrder.xlsm")
)

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlPersonnelEvents.bas"
$allowanceModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlPersonnelAllowanceRules.bas"
$orderExportModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlPersonnelEventOrderExport.bas"
$historyModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlPersonnelHistory.bas"
$staffLinkModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlStaffLinking.bas"
$legalActsModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlLegalActs.bas"
$enrollmentLinkModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlEnrollmentEventLink.bas"
$paymentRuleCatalogModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlPaymentRuleCatalog.bas"
$positionClassificationModulePath = Join-Path $workspace "CreateOrder.xlsm.modules\mdlPositionClassification.bas"
$testDirectory = Join-Path $workspace "_tmp_personnel_events_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_personnel_events_test.xlsm"

function Read-VbaText([string]$Path) {
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    try {
        $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
    }
    catch {
        $component = $Workbook.VBProject.VBComponents.Add(1)
        $component.Name = $ModuleName
    }
    $codeModule = $component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString($code)
}

function Add-ProbeModule($Workbook) {
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item("personnel_events_probe")) } catch {}
    $component = $Workbook.VBProject.VBComponents.Add(1)
    $component.Name = "personnel_events_probe"
    $component.CodeModule.AddFromString(@"
Option Explicit

Public Function ProbePersonnelEvents() As String
    Dim beforeState As Object, afterState As Object, payments As New Collection, payment As Object, payment2 As Object, payment3 As Object
    Dim transferID As String, exclusionID As String, correctionID As String, currentState As Object, correctionData As Object, events As Worksheet, rowNum As Long, transferCorrected As Boolean, correctionLinked As Boolean
    On Error GoTo Failure
    Set beforeState = CreateObject("Scripting.Dictionary")
    beforeState("fio") = "Test Employee"
    beforeState("personal_number") = "PE-001"
    beforeState("rank") = "Rank A"
    beforeState("position") = "Old position"
    beforeState("section") = "Old section"
    beforeState("military_unit") = "Unit A"
    beforeState("vus") = "100000A"
    beforeState("tariff_rank") = "4"
    beforeState("service_category") = "CONTRACT"
    beforeState("state_date") = DateSerial(2026, 7, 1)
    Set afterState = CreateObject("Scripting.Dictionary")
    afterState("fio") = "Test Employee"
    afterState("personal_number") = "PE-001"
    afterState("rank") = "Rank A"
    afterState("position") = "New position"
    afterState("section") = "New section"
    afterState("military_unit") = "Unit A"
    afterState("vus") = "200000B"
    afterState("tariff_rank") = "2"
    afterState("service_category") = "CONTRACT"
    afterState("state_date") = DateSerial(2026, 7, 2)
    Set payment = CreateObject("Scripting.Dictionary")
    payment("payment_type") = "Tariff allowance"
    payment("payment_code") = "TARIFF_1_4"
    payment("amount_kind") = "PERCENT"
    payment("amount_value") = "50"
    payment("start_date") = DateSerial(2026, 7, 2)
    payment("status") = "ACTIVE"
    payments.Add payment
    Set payment2 = CreateObject("Scripting.Dictionary")
    payment2("payment_type") = "FIZO"
    payment2("payment_code") = "FIZO_HIGH"
    payment2("amount_kind") = "PERCENT"
    payment2("amount_value") = "70"
    payment2("original_amount") = "70"
    payment2("applied_amount") = "70"
    payment2("cap_group") = "SPECIAL_ACHIEVEMENTS_P2"
    payment2("start_date") = DateSerial(2026, 7, 2)
    payment2("status") = "ACTIVE"
    payments.Add payment2
    Set payment3 = CreateObject("Scripting.Dictionary")
    payment3("payment_type") = "Driver"
    payment3("payment_code") = "DRIVER_C_D_CE"
    payment3("amount_kind") = "PERCENT"
    payment3("amount_value") = "30"
    payment3("original_amount") = "30"
    payment3("applied_amount") = "30"
    payment3("cap_group") = "SPECIAL_ACHIEVEMENTS_P2"
    payment3("start_date") = DateSerial(2026, 7, 2)
    payment3("status") = "ACTIVE"
    payments.Add payment3
    transferID = mdlPersonnelEvents.SaveTransferEvent("EMP-PE-001", DateSerial(2026, 7, 2), DateSerial(2026, 7, 2), beforeState, afterState, "Order 1", "Transfer basis", payments)
    Set currentState = mdlPersonnelEvents.GetCurrentPersonnelState("EMP-PE-001")
    If currentState("position") <> "New position" Then Err.Raise vbObjectError + 1, , "Current state was not updated"
    exclusionID = mdlPersonnelEvents.SaveExclusionEvent("EMP-PE-001", DateSerial(2026, 7, 10), DateSerial(2026, 7, 11), afterState, "Order 2", "Exclusion basis")
    Set correctionData = CreateObject("Scripting.Dictionary")
    correctionData("employee_id") = "EMP-PE-001"
    correctionData("event_type") = "TRANSFER"
    correctionData("event_date") = DateSerial(2026, 7, 12)
    correctionData("effective_date") = DateSerial(2026, 7, 12)
    correctionData("order_reference") = "Correction order"
    correctionData("basis_text") = "Correction basis"
    correctionData("corrects_event_id") = transferID
    correctionID = mdlPersonnelEvents.SavePersonnelEvent(correctionData, afterState, afterState)
    Set events = ThisWorkbook.Worksheets("PersonnelEvents")
    For rowNum = 2 To events.Cells(events.Rows.Count, 1).End(xlUp).Row
        If events.Cells(rowNum, 1).Value = transferID And events.Cells(rowNum, 6).Value = "CORRECTED" Then transferCorrected = True
        If events.Cells(rowNum, 1).Value = correctionID And events.Cells(rowNum, 14).Value = transferID Then correctionLinked = True
    Next rowNum
    If Not transferCorrected Or Not correctionLinked Then Err.Raise vbObjectError + 22, , "Correction did not preserve event history"
    If transferID = "" Or exclusionID = "" Then Err.Raise vbObjectError + 2, , "Event identifiers are blank"
    ProbePersonnelEvents = transferID & "|" & exclusionID
    Exit Function
Failure:
    ProbePersonnelEvents = "ERROR: " & Err.Description
End Function

Public Function ProbeAllowanceRules() As String
    Dim stateData As Object, ruleData As Object, results As Collection, item As Object
    Dim secondLevelState As Object, secondLevelResults As Collection
    Dim hasFizoExcluded As Boolean, hasTariff As Boolean, hasContract As Boolean, hasFixedActive As Boolean, hasSecondLevelFizo As Boolean, hasTimedMedal As Boolean
    On Error GoTo Failure
    Set stateData = CreateObject("Scripting.Dictionary")
    stateData("service_category") = "MOBILIZED"
    stateData("fizo_level") = "HIGH"
    stateData("sport_status") = "MASTER"
    stateData("tariff_rank") = "2"
    stateData("contract_430_eligible") = "YES"
    Set ruleData = CreateObject("Scripting.Dictionary")
    ruleData("mobilized_fixed_act_id") = "ACT-UNVERIFIED"
    Set results = mdlPersonnelAllowanceRules.EvaluatePersonnelAllowances(stateData, ruleData)
    For Each item In results
        Select Case item("payment_code")
            Case "FIZO": hasFizoExcluded = (item("status") = "NOT_APPLICABLE")
            Case "TARIFF_1_4": hasTariff = (item("status") = "ACTIVE" And item("amount_value") = "50")
            Case "MOBILIZATION_OR_SVO_CONTRACT": hasContract = (item("status") = "ACTIVE" And item("amount_value") = "60")
            Case "MOBILIZED_FIXED_158000": hasFixedActive = (item("status") = "ACTIVE" And item("amount_value") = "158000" And item("act_id") = mdlPersonnelEvents.LEGAL_ACT_UP_788)
        End Select
    Next item
    If Not hasFizoExcluded Or Not hasTariff Or Not hasContract Or Not hasFixedActive Then Err.Raise vbObjectError + 3, , "Allowance rules returned an unexpected result"

    Set secondLevelState = CreateObject("Scripting.Dictionary")
    secondLevelState("fizo_level") = "SECOND"
    Set secondLevelResults = mdlPersonnelAllowanceRules.EvaluatePersonnelAllowances(secondLevelState, CreateObject("Scripting.Dictionary"))
    For Each item In secondLevelResults
        If item("payment_code") = "FIZO_SECOND" Then hasSecondLevelFizo = (item("status") = "ACTIVE" And item("amount_value") = 15)
    Next item
    If Not hasSecondLevelFizo Then Err.Raise vbObjectError + 27, , "SECOND FIZO level must create a 15-percent assignment"

    Set secondLevelState = CreateObject("Scripting.Dictionary")
    secondLevelState("medal_code") = "COMBAT_DISTINCTION"
    Set ruleData = CreateObject("Scripting.Dictionary")
    ruleData("medal_award_date") = DateSerial(2026, 7, 14)
    ruleData("medal_award_document_reference") = "Order of award No. 12"
    Set secondLevelResults = mdlPersonnelAllowanceRules.EvaluatePersonnelAllowances(secondLevelState, ruleData)
    For Each item In secondLevelResults
        If item("payment_code") = "MEDAL_COMBAT_DISTINCTION" Then hasTimedMedal = (item("status") = "ACTIVE" And item("amount_value") = 30 And item("start_date") = DateSerial(2026, 7, 14) And item("end_date") = DateSerial(2027, 7, 13) And item("document_reference") = "Order of award No. 12")
    Next item
    If Not hasTimedMedal Then Err.Raise vbObjectError + 39, , "Medal payment must retain the 30-percent one-year award period"
    ProbeAllowanceRules = "OK"
    Exit Function
Failure:
    ProbeAllowanceRules = "ERROR: " & Err.Description
End Function

Public Function ProbePoint2Cap() As String
    Dim stateData As Object, ruleData As Object, results As Collection, item As Object
    Dim activeCount As Long, caps As Worksheet
    On Error GoTo Failure
    Set stateData = CreateObject("Scripting.Dictionary")
    stateData("fizo_level") = "HIGH"
    stateData("vus") = "310100"
    stateData("driver_c_d_ce") = "YES"
    Set ruleData = CreateObject("Scripting.Dictionary")
    Set results = mdlPersonnelAllowanceRules.EvaluatePersonnelAllowances(stateData, ruleData)
    For Each item In results
        If item("cap_group") = "SPECIAL_ACHIEVEMENTS_P2" And item("status") = "ACTIVE" Then activeCount = activeCount + 1
    Next item
    If activeCount <> 3 Then Err.Raise vbObjectError + 4, , "Point-2 cap overflow must keep every underlying rule active"
    Set caps = ThisWorkbook.Worksheets("PaymentCaps")
    If caps.Cells(2, 1).Value <> "SPECIAL_ACHIEVEMENTS_P2" Or caps.Cells(2, 2).Value <> 100 Then Err.Raise vbObjectError + 26, , "Confirmed point-2 cap was not stored in PaymentCaps"
    ProbePoint2Cap = "OK"
    Exit Function
Failure:
    ProbePoint2Cap = "ERROR: " & Err.Description
End Function

Public Function ProbePersonnelInputSheet() As String
    Dim ws As Worksheet, savedID As String, currentState As Object, assignments As Worksheet, events As Worksheet, rowNum As Long, hasTariff As Boolean, eventCountBeforeReset As Long
    On Error GoTo Failure
    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets("PersonnelEventInput")
    Set events = ThisWorkbook.Worksheets("PersonnelEvents")
    eventCountBeforeReset = events.Cells(events.Rows.Count, 1).End(xlUp).Row
    ws.Cells(4, 2).Value = "EXCLUSION"
    ws.Cells(22, 2).Value = "HIGH"
    ws.Cells(28, 2).Value = "EVT-OLD"
    ws.Cells(33, 2).Value = "EVT-CORRECTS"
    mdlPersonnelEvents.ResetPersonnelEventInput
    If ws.Cells(4, 2).Value <> "TRANSFER" Or ws.Cells(22, 2).Value <> "" Or ws.Cells(28, 2).Value <> "" Or ws.Cells(33, 2).Value <> "" Then Err.Raise vbObjectError + 33, , "Personnel input reset did not clear all event values"
    If events.Cells(events.Rows.Count, 1).End(xlUp).Row <> eventCountBeforeReset Then Err.Raise vbObjectError + 34, , "Personnel input reset must not create an event"
    ws.Cells(4, 2).Value = "TRANSFER"
    ws.Cells(5, 2).Value = "EMP-PE-001"
    ws.Cells(6, 2).Value = DateSerial(2026, 7, 12)
    ws.Cells(7, 2).Value = DateSerial(2026, 7, 12)
    ws.Cells(8, 2).Value = "Order 3"
    ws.Cells(9, 2).Value = "Input-sheet transfer"
    ws.Cells(12, 2).Value = "Input sheet position"
    ws.Cells(16, 2).Value = "2"
    ws.Cells(22, 2).Value = "HIGH"
    ws.Cells(23, 2).Value = "CMS"
    ws.Cells(24, 2).Value = "COMBAT_DISTINCTION"
    ws.Cells(25, 2).Value = "YES"
    ws.Cells(26, 2).Value = "YES"
    ws.Cells(42, 2).Value = DateSerial(2026, 7, 12)
    ws.Cells(43, 2).Value = "Award order No. 9"
    savedID = mdlPersonnelEvents.SavePersonnelEventInput(False)
    Set currentState = mdlPersonnelEvents.GetCurrentPersonnelState("EMP-PE-001")
    If savedID = "" Or currentState("position") <> "Input sheet position" Then Err.Raise vbObjectError + 5, , "Input-sheet event was not saved"
    If currentState("fizo_level") <> "HIGH" Or currentState("sport_status") <> "CMS" Or currentState("medal_code") <> "COMBAT_DISTINCTION" Or currentState("medal_award_date") <> DateSerial(2026, 7, 12) Or currentState("medal_award_document_reference") <> "Award order No. 9" Or currentState("driver_c_d_ce") <> "YES" Or currentState("contract_430_eligible") <> "YES" Then Err.Raise vbObjectError + 31, , "Allowance conditions were not persisted in current state"
    mdlPersonnelEvents.PreparePersonnelEventCorrection "EMP-PE-001", savedID
    If ws.Cells(22, 2).Value <> "HIGH" Or ws.Cells(23, 2).Value <> "CMS" Or ws.Cells(24, 2).Value <> "COMBAT_DISTINCTION" Or ws.Cells(25, 2).Value <> "YES" Or ws.Cells(26, 2).Value <> "YES" Or ws.Cells(42, 2).Value <> DateSerial(2026, 7, 12) Or ws.Cells(43, 2).Value <> "Award order No. 9" Then Err.Raise vbObjectError + 32, , "Correction form did not preload persisted allowance conditions"
    Set assignments = ThisWorkbook.Worksheets("PaymentAssignments")
    For rowNum = 2 To assignments.Cells(assignments.Rows.Count, 1).End(xlUp).Row
        If assignments.Cells(rowNum, 3).Value = savedID And assignments.Cells(rowNum, 5).Value = "TARIFF_1_4" Then hasTariff = True
        If assignments.Cells(rowNum, 3).Value = savedID And assignments.Cells(rowNum, 5).Value = "MEDAL_COMBAT_DISTINCTION" Then
            If assignments.Cells(rowNum, 9).Value <> DateSerial(2026, 7, 12) Or assignments.Cells(rowNum, 10).Value <> DateSerial(2027, 7, 11) Or assignments.Cells(rowNum, 16).Value <> "Award order No. 9" Then Err.Raise vbObjectError + 40, , "Saved medal assignment does not retain its one-year period or award document"
        End If
    Next rowNum
    If Not hasTariff Then Err.Raise vbObjectError + 7, , "Input-sheet transfer did not create allowance assignments"
    ProbePersonnelInputSheet = "OK"
    Exit Function
Failure:
    ProbePersonnelInputSheet = "ERROR: " & Err.Description
End Function

Public Function ProbePersonnelOrderExport() As String
    Dim eventID As String, stateData As Object, payments As New Collection, p1 As Object, p2 As Object, p3 As Object, events As Worksheet, rowNum As Long, outputPath As String
    On Error GoTo Failure
    Set stateData = mdlPersonnelEvents.GetCurrentPersonnelState("EMP-PE-001")
    Set p1 = CreateObject("Scripting.Dictionary")
    p1("payment_type") = "FIZO"
    p1("payment_code") = "FIZO_HIGH"
    p1("amount_kind") = "PERCENT"
    p1("amount_value") = "70"
    p1("original_amount") = "70"
    p1("applied_amount") = "70"
    p1("cap_group") = "SPECIAL_ACHIEVEMENTS_P2"
    p1("status") = "ACTIVE"
    p1("act_id") = mdlPersonnelEvents.LEGAL_ACT_MO_430
    payments.Add p1
    Set p2 = CreateObject("Scripting.Dictionary")
    p2("payment_type") = "VUS"
    p2("payment_code") = "VUS_310100_310101"
    p2("amount_kind") = "PERCENT"
    p2("amount_value") = "50"
    p2("original_amount") = "50"
    p2("applied_amount") = "50"
    p2("cap_group") = "SPECIAL_ACHIEVEMENTS_P2"
    p2("status") = "ACTIVE"
    p2("act_id") = mdlPersonnelEvents.LEGAL_ACT_MO_430
    payments.Add p2
    Set p3 = CreateObject("Scripting.Dictionary")
    p3("payment_type") = "Mobilized social payment"
    p3("payment_code") = "MOBILIZED_FIXED_158000"
    p3("amount_kind") = "FIXED_AMOUNT"
    p3("amount_value") = "158000"
    p3("original_amount") = "158000"
    p3("applied_amount") = "158000"
    p3("cap_group") = "SEPARATE_LEGAL_ACT"
    p3("status") = "ACTIVE"
    p3("act_id") = mdlPersonnelEvents.LEGAL_ACT_UP_788
    payments.Add p3
    eventID = mdlPersonnelEvents.SaveTransferEvent("EMP-PE-001", DateSerial(2026, 7, 13), DateSerial(2026, 7, 13), stateData, stateData, "Order cap", "Cap basis", payments)
    outputPath = mdlPersonnelEventOrderExport.ExportPersonnelEventOrder(eventID)
    Set events = ThisWorkbook.Worksheets("PersonnelEvents")
    For rowNum = 2 To events.Cells(events.Rows.Count, 1).End(xlUp).Row
        If events.Cells(rowNum, 1).Value = eventID Then Exit For
    Next rowNum
    If events.Cells(rowNum, 6).Value <> "EXPORTED" Then Err.Raise vbObjectError + 21, , "Personnel event status was not updated after Word export"
    ProbePersonnelOrderExport = outputPath
    Exit Function
Failure:
    ProbePersonnelOrderExport = "ERROR: " & Err.Description
End Function

Public Function ProbePersonnelHistory() As String
    Dim ws As Worksheet, inputSheet As Worksheet, eventSheet As Worksheet
    Dim selectedEventID As String
    Dim eventCountBefore As Long, eventCountAfter As Long
    On Error GoTo Failure
    mdlPersonnelHistory.OpenPersonnelHistory
    Set ws = ThisWorkbook.Worksheets("PersonnelHistory")
    ws.Cells(3, 2).Value = "PE-001"
    mdlPersonnelHistory.SearchPersonnelHistory
    If ws.Cells(4, 2).Value <> "EMP-PE-001" Then Err.Raise vbObjectError + 8, , "History search did not resolve the employee"
    If Application.WorksheetFunction.CountIf(ws.UsedRange, "EVT-*") < 3 Then Err.Raise vbObjectError + 9, , "History did not show saved events"
    If Application.WorksheetFunction.CountIf(ws.UsedRange, "TERMINATED") < 1 Then Err.Raise vbObjectError + 10, , "History did not show terminated payment"
    If Application.WorksheetFunction.CountIf(ws.UsedRange, "PERSONNEL_ORDER") < 1 Then Err.Raise vbObjectError + 11, , "History did not show generated document"
    If Application.WorksheetFunction.CountIf(ws.UsedRange, "SYNCED") < 1 Then Err.Raise vbObjectError + 31, , "History did not show staff state synchronization audit"
    selectedEventID = CStr(ws.Cells(16, 1).Value)
    If Left(selectedEventID, 4) <> "EVT-" Then Err.Raise vbObjectError + 28, , "History did not expose an EventID for correction preparation"
    Set eventSheet = ThisWorkbook.Worksheets("PersonnelEvents")
    eventCountBefore = eventSheet.Cells(eventSheet.Rows.Count, 1).End(xlUp).Row
    ws.Cells(16, 1).Select
    mdlPersonnelHistory.PreparePersonnelHistoryCorrection False
    Set inputSheet = ThisWorkbook.Worksheets("PersonnelEventInput")
    eventCountAfter = eventSheet.Cells(eventSheet.Rows.Count, 1).End(xlUp).Row
    If inputSheet.Cells(5, 2).Value <> "EMP-PE-001" Or inputSheet.Cells(33, 2).Value <> selectedEventID Then Err.Raise vbObjectError + 29, , "Correction form was not prepared from the selected history event"
    If eventCountBefore <> eventCountAfter Then Err.Raise vbObjectError + 30, , "Preparing a correction must not save an event"
    ProbePersonnelHistory = "OK"
    Exit Function
Failure:
    ProbePersonnelHistory = "ERROR: " & Err.Description
End Function

Public Function ProbeStaffLinking() As String
    Dim staff As Worksheet, review As Worksheet, employees As Worksheet, syncLog As Worksheet, currentState As Object
    Dim personalColumn As Long, rankColumn As Long, fioColumn As Long, positionColumn As Long, unitColumn As Long
    Dim staffRow As Long, reviewRow As Long, employeeRow As Long, eventCount As Long, assignmentCount As Long
    On Error GoTo Failure
    Set staff = GetStaffWorksheet()
    If Not FindColumnNumbers(staff, personalColumn, rankColumn, fioColumn, positionColumn, unitColumn) Then Err.Raise vbObjectError + 12, , "Staff columns were not found"
    staffRow = staff.Cells(staff.Rows.Count, personalColumn).End(xlUp).Row + 1
    staff.Cells(staffRow, personalColumn).Value = "PE-001"
    staff.Cells(staffRow, fioColumn).Value = "Test Employee"
    staff.Cells(staffRow, rankColumn).Value = "Linked test rank"
    staff.Cells(staffRow, positionColumn).Value = "Linked test position"
    staff.Cells(staffRow, unitColumn).Value = "Linked test unit"
    eventCount = ThisWorkbook.Worksheets("PersonnelEvents").Cells(ThisWorkbook.Worksheets("PersonnelEvents").Rows.Count, 1).End(xlUp).Row
    assignmentCount = ThisWorkbook.Worksheets("PaymentAssignments").Cells(ThisWorkbook.Worksheets("PaymentAssignments").Rows.Count, 1).End(xlUp).Row
    mdlStaffLinking.BuildStaffLinkCandidates False
    Set review = ThisWorkbook.Worksheets("StaffLinkReview")
    For reviewRow = 5 To review.Cells(review.Rows.Count, 1).End(xlUp).Row
        If review.Cells(reviewRow, 1).Value = "EMP-PE-001" Then Exit For
    Next reviewRow
    If review.Cells(reviewRow, 4).Value <> "CANDIDATE" Then Err.Raise vbObjectError + 13, , "Staff candidate was not created"
    review.Cells(reviewRow, 9).Value = "CONFIRM"
    mdlStaffLinking.ConfirmStaffLinkSelections False
    Set employees = ThisWorkbook.Worksheets("Employees")
    For employeeRow = 2 To employees.Cells(employees.Rows.Count, 1).End(xlUp).Row
        If employees.Cells(employeeRow, 1).Value = "EMP-PE-001" Then Exit For
    Next employeeRow
    If employees.Cells(employeeRow, 6).Value <> "LINKED" Or employees.Cells(employeeRow, 7).Value <> "STAFF_ROW:" & CStr(staffRow) Then Err.Raise vbObjectError + 14, , "Staff link was not confirmed"
    review.Cells(reviewRow, 9).Value = "SYNC"
    mdlStaffLinking.SyncConfirmedStaffState False
    Set currentState = mdlPersonnelEvents.GetCurrentPersonnelState("EMP-PE-001")
    If currentState("rank") <> "Linked test rank" Or currentState("position") <> "Linked test position" Or currentState("military_unit") <> "Linked test unit" Then Err.Raise vbObjectError + 35, , "Linked staff synchronization did not update current state"
    Set syncLog = ThisWorkbook.Worksheets("StaffStateSyncLog")
    If syncLog.Cells(syncLog.Rows.Count, 1).End(xlUp).Row < 2 Or syncLog.Cells(syncLog.Rows.Count, 4).End(xlUp).Value <> "SYNCED" Then Err.Raise vbObjectError + 36, , "Linked staff synchronization was not audited"
    If eventCount <> ThisWorkbook.Worksheets("PersonnelEvents").Cells(ThisWorkbook.Worksheets("PersonnelEvents").Rows.Count, 1).End(xlUp).Row Then Err.Raise vbObjectError + 15, , "Linking changed event history"
    If assignmentCount <> ThisWorkbook.Worksheets("PaymentAssignments").Cells(ThisWorkbook.Worksheets("PaymentAssignments").Rows.Count, 1).End(xlUp).Row Then Err.Raise vbObjectError + 16, , "Linking changed payment history"
    ProbeStaffLinking = "OK"
    Exit Function
Failure:
    ProbeStaffLinking = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentInput() As String
    Dim ws As Worksheet, eventID As String, employeeID As String, stateData As Object, employees As Worksheet
    Dim employeeCount As Long, eventCount As Long
    On Error GoTo Failure
    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets("PersonnelEventInput")
    Set employees = ThisWorkbook.Worksheets("Employees")
    employeeCount = employees.Cells(employees.Rows.Count, 1).End(xlUp).Row
    eventCount = ThisWorkbook.Worksheets("PersonnelEvents").Cells(ThisWorkbook.Worksheets("PersonnelEvents").Rows.Count, 1).End(xlUp).Row
    ws.Cells(4, 2).Value = "ENROLLMENT"
    ws.Cells(5, 2).ClearContents
    ws.Cells(6, 2).Value = DateSerial(2026, 7, 14)
    ws.Cells(7, 2).Value = DateSerial(2026, 7, 14)
    ws.Cells(8, 2).Value = "Enrollment order"
    ws.Cells(9, 2).Value = "Enrollment basis"
    ws.Cells(11, 2).Value = "Rank E"
    ws.Cells(12, 2).Value = "Enrollment position"
    ws.Cells(16, 2).Value = "2"
    ws.Cells(19, 2).Value = "CONTRACT"
    ws.Cells(29, 2).Value = "Enrollment Employee"
    ws.Cells(30, 2).Value = "ENR-001"
    ws.Cells(31, 2).Value = "TN-001"
    ws.Cells(32, 2).Value = "MANUAL"
    ws.Cells(33, 2).ClearContents
    ws.Range(ws.Cells(22, 2), ws.Cells(27, 2)).ClearContents
    eventID = mdlPersonnelEvents.SavePersonnelEventInput(False)
    employeeID = ws.Cells(5, 2).Value
    Set stateData = mdlPersonnelEvents.GetCurrentPersonnelState(employeeID)
    If Left(employeeID, 4) <> "EMP-" Or eventID = "" Then Err.Raise vbObjectError + 17, , "Enrollment did not create an internal employee ID"
    If stateData("position") <> "Enrollment position" Or stateData("service_category") <> "CONTRACT" Then Err.Raise vbObjectError + 18, , "Enrollment did not save current state"
    If employees.Cells(employees.Rows.Count, 1).End(xlUp).Row <> employeeCount + 1 Then Err.Raise vbObjectError + 19, , "Enrollment did not create one employee card"
    If ThisWorkbook.Worksheets("PersonnelEvents").Cells(ThisWorkbook.Worksheets("PersonnelEvents").Rows.Count, 1).End(xlUp).Row <> eventCount + 1 Then Err.Raise vbObjectError + 20, , "Enrollment did not create one event"
    ProbeEnrollmentInput = "OK"
    Exit Function
Failure:
    ProbeEnrollmentInput = "ERROR: " & Err.Description
End Function

Public Function ProbeLegalActInput() As String
    Dim ws As Worksheet, acts As Worksheet, actID As String, rowNum As Long, found As Boolean
    On Error GoTo Failure
    mdlLegalActs.OpenLegalActInput
    Set ws = ThisWorkbook.Worksheets("LegalActInput")
    ws.Cells(4, 2).ClearContents
    ws.Cells(5, 2).Value = "ORDER"
    ws.Cells(6, 2).Value = "TEST-001"
    ws.Cells(7, 2).Value = DateSerial(2026, 7, 13)
    ws.Cells(8, 2).Value = "Test legal act"
    ws.Cells(9, 2).Value = "1"
    ws.Cells(10, 2).Value = DateSerial(2026, 7, 13)
    ws.Cells(11, 2).Value = DateSerial(2026, 12, 31)
    ws.Cells(12, 2).Value = "TEST"
    ws.Cells(13, 2).Value = "Acceptance test"
    actID = mdlLegalActs.SaveLegalActInput(False)
    Set acts = ThisWorkbook.Worksheets("LegalActs")
    For rowNum = 2 To acts.Cells(acts.Rows.Count, 1).End(xlUp).Row
        If acts.Cells(rowNum, 1).Value = actID And acts.Cells(rowNum, 3).Value = "TEST-001" Then found = True
    Next rowNum
    If Left(actID, 4) <> "ACT-" Or Not found Then Err.Raise vbObjectError + 23, , "Legal act was not saved"
    ProbeLegalActInput = "OK"
    Exit Function
Failure:
    ProbeLegalActInput = "ERROR: " & Err.Description
End Function

Public Function ProbeEnrollmentDocumentLink() As String
    Dim inputSheet As Worksheet, documents As Worksheet, events As Worksheet
    Dim eventID As String, filePath As String, documentID As String, rowNum As Long, found As Boolean, exported As Boolean
    On Error GoTo Failure
    Set inputSheet = ThisWorkbook.Worksheets("PersonnelEventInput")
    eventID = inputSheet.Cells(28, 2).Value
    Set documents = ThisWorkbook.Worksheets("DocumentRegistry")
    For rowNum = 2 To documents.Cells(documents.Rows.Count, 1).End(xlUp).Row
        If documents.Cells(rowNum, 3).Value = "PERSONNEL_ORDER" Then filePath = documents.Cells(rowNum, 6).Value
    Next rowNum
    documentID = mdlEnrollmentEventLink.RegisterEnrollmentOrderForEvent(eventID, filePath, "Enrollment order")
    For rowNum = 2 To documents.Cells(documents.Rows.Count, 1).End(xlUp).Row
        If documents.Cells(rowNum, 1).Value = documentID And documents.Cells(rowNum, 2).Value = eventID And documents.Cells(rowNum, 3).Value = "ENROLLMENT_ORDER" Then found = True
    Next rowNum
    Set events = ThisWorkbook.Worksheets("PersonnelEvents")
    For rowNum = 2 To events.Cells(events.Rows.Count, 1).End(xlUp).Row
        If events.Cells(rowNum, 1).Value = eventID And events.Cells(rowNum, 6).Value = "EXPORTED" Then exported = True
    Next rowNum
    If Not found Or Not exported Then Err.Raise vbObjectError + 24, , "Enrollment document was not linked"
    ProbeEnrollmentDocumentLink = "OK"
    Exit Function
Failure:
    ProbeEnrollmentDocumentLink = "ERROR: " & Err.Description
End Function

Public Function ProbePaymentRuleCatalog() As String
    Dim ws As Worksheet, rules As Worksheet, ruleID As String, rowNum As Long, found As Boolean
    On Error GoTo Failure
    mdlPaymentRuleCatalog.OpenPaymentRuleInput
    Set ws = ThisWorkbook.Worksheets("PaymentRuleInput")
    ws.Cells(4, 2).ClearContents
    ws.Cells(5, 2).Value = "TEST_PAYMENT"
    ws.Cells(6, 2).Value = "TEST_BASIS"
    ws.Cells(7, 2).Value = "PERCENT"
    ws.Cells(8, 2).Value = "10"
    ws.Cells(9, 2).Value = "FIELD"
    ws.Cells(10, 2).Value = "EQUALS"
    ws.Cells(11, 2).Value = "YES"
    ws.Cells(23, 2).Value = DateSerial(2026, 7, 13)
    ws.Cells(24, 2).Value = DateSerial(2026, 12, 31)
    ruleID = mdlPaymentRuleCatalog.SavePaymentRuleInput(False)
    Set rules = ThisWorkbook.Worksheets("PaymentRules")
    For rowNum = 2 To rules.Cells(rules.Rows.Count, 1).End(xlUp).Row
        If rules.Cells(rowNum, 1).Value = ruleID And rules.Cells(rowNum, 2).Value = "TEST_PAYMENT" And rules.Cells(rowNum, 22).Value = "DRAFT" Then found = True
    Next rowNum
    If Left(ruleID, 5) <> "RULE-" Or Not found Then Err.Raise vbObjectError + 25, , "Payment rule catalog row was not saved"
    ProbePaymentRuleCatalog = "OK"
    Exit Function
Failure:
    ProbePaymentRuleCatalog = "ERROR: " & Err.Description
End Function

Public Function ProbePositionClassification() As String
    Dim ws As Worksheet, dataSheet As Worksheet, recordID As String, rowNum As Long, found As Boolean
    On Error GoTo Failure
    mdlPositionClassification.OpenPositionClassificationInput
    Set ws = ThisWorkbook.Worksheets("PositionClassificationInput")
    ws.Cells(4, 2).ClearContents
    ws.Cells(5, 2).Value = "TEST_POSITION"
    ws.Cells(6, 2).Value = "TEST-STAFF"
    ws.Cells(7, 2).Value = "Test position"
    ws.Cells(8, 2).Value = "DRIVER"
    ws.Cells(11, 2).Value = "TEST"
    recordID = mdlPositionClassification.SavePositionClassificationInput(False)
    Set dataSheet = ThisWorkbook.Worksheets("PositionClassification")
    For rowNum = 2 To dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row
        If dataSheet.Cells(rowNum, 1).Value = recordID And dataSheet.Cells(rowNum, 5).Value = "DRIVER" And dataSheet.Cells(rowNum, 9).Value = "DRAFT" Then found = True
    Next rowNum
    If Left(recordID, 4) <> "POS-" Or Not found Then Err.Raise vbObjectError + 27, , "Position classification was not saved"
    ProbePositionClassification = "OK"
    Exit Function
Failure:
    ProbePositionClassification = "ERROR: " & Err.Description
End Function
"@)
}

if (Test-Path $testDirectory) { Remove-Item -LiteralPath $testDirectory -Recurse -Force }
New-Item -ItemType Directory -Path $testDirectory | Out-Null
Copy-Item -LiteralPath $WorkbookPath -Destination $testWorkbookPath

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlPersonnelEvents" -ModulePath $modulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlPersonnelAllowanceRules" -ModulePath $allowanceModulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlPersonnelEventOrderExport" -ModulePath $orderExportModulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlPersonnelHistory" -ModulePath $historyModulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlStaffLinking" -ModulePath $staffLinkModulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlLegalActs" -ModulePath $legalActsModulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlEnrollmentEventLink" -ModulePath $enrollmentLinkModulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlPaymentRuleCatalog" -ModulePath $paymentRuleCatalogModulePath
    Import-CodeModuleText -Workbook $workbook -ModuleName "mdlPositionClassification" -ModulePath $positionClassificationModulePath
    Add-ProbeModule -Workbook $workbook
    $result = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbePersonnelEvents")
    if ($result -like "ERROR:*") { throw $result }
    $events = $workbook.Worksheets("PersonnelEvents")
    $assignments = $workbook.Worksheets("PaymentAssignments")
    if ($events.Cells(3, 3).Value2 -ne "EXCLUSION") { throw "Exclusion event was not persisted." }
    if ($assignments.Cells(2, 11).Value2 -ne "TERMINATED") { throw "Active payment was not terminated by exclusion." }
    $allowanceResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbeAllowanceRules")
    if ($allowanceResult -ne "OK") { throw "Allowance rule test failed: $allowanceResult" }
    $capResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbePoint2Cap")
    if ($capResult -ne "OK") { throw "Point-2 cap test failed: $capResult" }
    $inputResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbePersonnelInputSheet")
    if ($inputResult -ne "OK") { throw "Personnel input sheet test failed: $inputResult" }
    $orderPath = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbePersonnelOrderExport")
    if ($orderPath -like "ERROR:*" -or -not (Test-Path -LiteralPath $orderPath)) { throw "Personnel order export failed: $orderPath" }
    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $doc = $word.Documents.Open($orderPath, $false, $true)
        $wordText = [string]$doc.Content.Text
        if ($wordText -notlike "*100*" -or $wordText -notlike "*70%*" -or $wordText -notlike "*50%*" -or $wordText -notlike "*430*" -or $wordText -notlike "*788*" -or $wordText -notlike "*158000*") { throw "Personnel order did not retain the confirmed legal headings, point-2 cap and all original grounds." }
    }
    finally {
        if ($doc -ne $null) { $doc.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($word -ne $null) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    }
    $enrollmentResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbeEnrollmentInput")
    if ($enrollmentResult -ne "OK") { throw "Enrollment-input test failed: $enrollmentResult" }
    $legalActResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbeLegalActInput")
    if ($legalActResult -ne "OK") { throw "Legal-act test failed: $legalActResult" }
    $enrollmentDocumentResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbeEnrollmentDocumentLink")
    if ($enrollmentDocumentResult -ne "OK") { throw "Enrollment-document link test failed: $enrollmentDocumentResult" }
    $paymentRuleResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbePaymentRuleCatalog")
    if ($paymentRuleResult -ne "OK") { throw "Payment-rule catalog test failed: $paymentRuleResult" }
    $positionClassificationResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbePositionClassification")
    if ($positionClassificationResult -ne "OK") { throw "Position-classification test failed: $positionClassificationResult" }
    $staffLinkResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbeStaffLinking")
    if ($staffLinkResult -ne "OK") { throw "Staff-link test failed: $staffLinkResult" }
    $historyResult = [string]$excel.Run("'$($workbook.Name)'!personnel_events_probe.ProbePersonnelHistory")
    if ($historyResult -ne "OK") { throw "Personnel history test failed: $historyResult" }
    Write-Output "Personnel event acceptance passed: $result"
}
finally {
    if ($workbook -ne $null) { try { $workbook.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    if ($workbook -ne $null) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {} }
    if ($excel -ne $null) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
}
