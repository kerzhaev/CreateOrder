Attribute VB_Name = "mdlPersonnelEvents"
Option Explicit

' Persistent personnel-event ledger. This module deliberately stores data in
' append-only service sheets so the history does not depend on later updates
' of the Staff sheet.

Private Const SHEET_EMPLOYEES As String = "Employees"
Private Const SHEET_CURRENT_STATE As String = "EmployeeCurrentState"
Private Const SHEET_EVENTS As String = "PersonnelEvents"
Private Const SHEET_SNAPSHOTS As String = "PersonnelStateSnapshots"
Private Const SHEET_ASSIGNMENTS As String = "PaymentAssignments"
Private Const SHEET_DOCUMENTS As String = "DocumentRegistry"
Private Const SHEET_LEGAL_ACTS As String = "LegalActs"
Private Const SHEET_EVENT_INPUT As String = "PersonnelEventInput"

Public Const EVENT_TYPE_ENROLLMENT As String = "ENROLLMENT"
Public Const EVENT_TYPE_TRANSFER As String = "TRANSFER"
Public Const EVENT_TYPE_EXCLUSION As String = "EXCLUSION"

Public Const EVENT_STATUS_DRAFT As String = "DRAFT"
Public Const EVENT_STATUS_VERIFIED As String = "VERIFIED"
Public Const EVENT_STATUS_EXPORTED As String = "EXPORTED"
Public Const EVENT_STATUS_CANCELLED As String = "CANCELLED"
Public Const EVENT_STATUS_CORRECTED As String = "CORRECTED"
Public Const LEGAL_ACT_MO_727 As String = "MO-727-20191206"
Public Const LEGAL_ACT_MO_430 As String = "MO-430-DSP-20190731"
Public Const LEGAL_ACT_UP_788 As String = "UP-788-20221102"
Public Const LEGAL_ACT_MO_780 As String = "MO-780-20221219"

Public Sub EnsurePersonnelEventInfrastructure()
    EnsureServiceSheet SHEET_EMPLOYEES, Array("EmployeeID", "FIO", "PersonalNumber", "TableNumber", "SourceMode", "StaffLinkStatus", "StaffReference", "CreatedAt", "UpdatedAt", "IsActive")
    EnsureServiceSheet SHEET_CURRENT_STATE, Array("EmployeeID", "Rank", "RankEffectiveDate", "Position", "Section", "MilitaryUnit", "VUS", "TariffRank", "PositionSalary", "RankSalary", "ServiceCategory", "ContractKind", "ContractBasis", "StateDate", "SourceEventID", "LastEventID", "FizoLevel", "SportStatus", "MedalCode", "DriverCDCE", "Contract430Eligible", "MedalAwardDate", "MedalAwardDocumentReference")
    EnsureServiceSheet SHEET_EVENTS, Array("EventID", "EmployeeID", "EventType", "EventDate", "EffectiveDate", "Status", "BeforeSnapshotID", "AfterSnapshotID", "OrderReference", "BasisText", "OperatorName", "CreatedAt", "UpdatedAt", "CorrectsEventID", "Comment", "HandoverDate", "AcceptanceDate", "DutyStartDate", "DestinationUnit", "DestinationLocation", "MaterialAssistanceStatus", "MainLeaveStatus", "AdditionalLeaveStatus")
    EnsureServiceSheet SHEET_SNAPSHOTS, Array("SnapshotID", "EventID", "SnapshotKind", "EmployeeID", "Rank", "Position", "Section", "MilitaryUnit", "VUS", "TariffRank", "PositionSalary", "RankSalary", "ServiceCategory", "ContractKind", "ContractBasis", "StateDate", "SerializedData", "CreatedAt")
    EnsureServiceSheet SHEET_ASSIGNMENTS, Array("AssignmentID", "EmployeeID", "EventID", "PaymentType", "PaymentCode", "AmountKind", "AmountValue", "CalculationBase", "StartDate", "EndDate", "Status", "TerminationEventID", "ActID", "ActPoint", "FactualBasis", "DocumentReference", "OriginalAmount", "AppliedAmount", "CapGroup", "Explanation", "CreatedAt", "UpdatedAt")
    EnsureServiceSheet SHEET_DOCUMENTS, Array("DocumentID", "EventID", "DocumentType", "DocumentNumber", "DocumentDate", "FilePath", "TemplateName", "TemplateVersion", "FileHash", "Status", "LastError", "CreatedAt")
    EnsureServiceSheet SHEET_LEGAL_ACTS, Array("ActID", "ActType", "ActNumber", "ActDate", "Title", "Revision", "EffectiveFrom", "EffectiveTo", "AccessMark", "Note", "CreatedAt", "UpdatedAt")
    EnsureServiceSheet "PaymentRules", Array("RuleID", "PaymentCode", "BasisCode", "AmountKind", "AmountValue", "ConditionType", "ConditionOperator", "ExpectedValue", "FactSource", "RequiredDocuments", "StartDateSource", "EndRule", "ActID", "ActPoint", "CapGroup", "Priority", "Severity", "ExplanationTemplate", "WordTemplate", "EffectiveFrom", "EffectiveTo", "RuleStatus", "CreatedAt", "UpdatedAt")
    EnsureServiceSheet "PaymentCaps", Array("CapGroup", "MaxPercent", "DistributionMethod", "PriorityPolicy", "EffectiveFrom", "EffectiveTo", "ActID", "ActPoint", "Status", "Note", "CreatedAt", "UpdatedAt")
    EnsureServiceSheet "PositionClassification", Array("ClassificationID", "PositionKey", "StaffCode", "PositionText", "GroupCode", "CommandLevel", "OtherFlags", "Source", "ReviewStatus", "Note", "CreatedAt", "UpdatedAt")
    EnsureConfirmedPoint2Cap
    EnsureConfirmedPersonnelLegalActs
    EnsurePersonnelEventInputSheet
End Sub

Public Sub EnsureConfirmedPersonnelLegalActs()
    EnsureLegalAct LEGAL_ACT_MO_727, "ORDER", "727", DateSerial(2019, 12, 6), "Об определении Порядка обеспечения денежным довольствием военнослужащих Вооруженных Сил Российской Федерации и предоставления им и членам их семей отдельных выплат", DateSerial(2020, 1, 27), "Подтвержденное основание для групп Word-проекта."
    EnsureLegalAct LEGAL_ACT_MO_430, "ORDER", "430дсп", DateSerial(2019, 7, 31), "Об утверждении Правил выплаты ежемесячной надбавки за особые достижения в службе военнослужащим Вооруженных Сил Российской Федерации, проходящим военную службу по контракту", DateSerial(2019, 7, 31), "Подтвержденное основание особых достижений."
    EnsureLegalAct LEGAL_ACT_UP_788, "DECREE", "788", DateSerial(2022, 11, 2), "О ежемесячной социальной выплате гражданам Российской Федерации, призванным на военную службу по мобилизации в Вооруженные Силы Российской Федерации", DateSerial(2022, 11, 2), "Размер 158000 рублей; порядок определен приказом МО № 780."
    EnsureLegalAct LEGAL_ACT_MO_780, "ORDER", "780", DateSerial(2022, 12, 19), "Об определении Порядка осуществления ежемесячной социальной выплаты гражданам Российской Федерации, призванным на военную службу по мобилизации в Вооруженные Силы Российской Федерации", DateSerial(2022, 9, 21), "Осуществление выплаты с 21.09.2022."
End Sub

Private Sub EnsureLegalAct(ByVal actID As String, ByVal actType As String, ByVal actNumber As String, ByVal actDate As Date, ByVal titleText As String, ByVal effectiveFrom As Date, ByVal noteText As String)
    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_LEGAL_ACTS)
    For rowNum = 2 To LastDataRow(ws)
        If StrComp(SafeText(ws.Cells(rowNum, 1).Value), actID, vbTextCompare) = 0 Then Exit Sub
    Next rowNum
    rowNum = NextDataRow(ws)
    ws.Cells(rowNum, 1).Value = actID
    ws.Cells(rowNum, 2).Value = actType
    ws.Cells(rowNum, 3).Value = actNumber
    ws.Cells(rowNum, 4).Value = actDate
    ws.Cells(rowNum, 5).Value = titleText
    ws.Cells(rowNum, 7).Value = effectiveFrom
    ws.Cells(rowNum, 9).Value = "CONFIRMED"
    ws.Cells(rowNum, 10).Value = noteText
    ws.Cells(rowNum, 11).Value = Now
    ws.Cells(rowNum, 12).Value = Now
End Sub

Private Sub EnsureConfirmedPoint2Cap()
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("PaymentCaps")
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = "SPECIAL_ACHIEVEMENTS_P2" Then Exit Sub
    Next rowNum
    rowNum = NextDataRow(ws)
    ws.Cells(rowNum, 1).Value = "SPECIAL_ACHIEVEMENTS_P2"
    ws.Cells(rowNum, 2).Value = 100
    ws.Cells(rowNum, 3).Value = "GROUP_CAP"
    ws.Cells(rowNum, 4).Value = "RETAIN_ORIGINAL_GROUNDS"
    ws.Cells(rowNum, 9).Value = "CONFIRMED"
    ws.Cells(rowNum, 10).Value = "Confirmed project rule for point-2 total."
    ws.Cells(rowNum, 11).Value = Now
    ws.Cells(rowNum, 12).Value = Now
End Sub

Public Sub OpenPersonnelEventInput()
    EnsurePersonnelEventInfrastructure
    ThisWorkbook.Worksheets(SHEET_EVENT_INPUT).Activate
End Sub

Public Sub OpenPersonnelActionMenu()
    EnsurePersonnelEventInfrastructure
    frmPersonnelActionWizard.ShowActionMenu
End Sub

Public Sub OpenPersonnelTransferAction()
    OpenPersonnelActionWizard EVENT_TYPE_TRANSFER
End Sub

Public Sub OpenPersonnelExclusionAction()
    OpenPersonnelActionWizard EVENT_TYPE_EXCLUSION
End Sub

Public Sub OpenPersonnelActionWizard(ByVal eventType As String)
    Dim normalizedType As String

    normalizedType = UCase$(Trim$(eventType))
    If normalizedType <> EVENT_TYPE_TRANSFER And normalizedType <> EVENT_TYPE_EXCLUSION Then
        Err.Raise vbObjectError + 660, "mdlPersonnelEvents", "Personnel action wizard supports TRANSFER and EXCLUSION only."
    End If
    PrepareNewPersonnelAction normalizedType
    frmPersonnelActionWizard.Show
End Sub

Public Function GetPersonnelWizardValue(ByVal fieldKey As String) As Variant
    EnsurePersonnelEventInfrastructure
    GetPersonnelWizardValue = GetInputValue(ThisWorkbook.Worksheets(SHEET_EVENT_INPUT), fieldKey)
End Function

Public Sub SetPersonnelWizardValue(ByVal fieldKey As String, ByVal fieldValue As Variant)
    EnsurePersonnelEventInfrastructure
    SetInputValue ThisWorkbook.Worksheets(SHEET_EVENT_INPUT), fieldKey, fieldValue
End Sub

Public Function LoadPersonnelWizardCurrentState() As Boolean
    Dim ws As Worksheet
    Dim employeeID As String
    Dim stateData As Object

    EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    employeeID = SafeText(GetInputValue(ws, "employee_id"))
    If employeeID = "" Then
        Application.StatusBar = "Enter EmployeeID before loading current state."
        Exit Function
    End If
    Set stateData = GetCurrentPersonnelState(employeeID)
    If stateData.Count = 0 Then
        Application.StatusBar = "Current state was not found for this EmployeeID."
        Exit Function
    End If
    PopulateCurrentStateInput ws, stateData
    Application.StatusBar = "Current state loaded into the personnel action form."
    LoadPersonnelWizardCurrentState = True
End Function

Public Function SavePersonnelWizardAction() As String
    SavePersonnelWizardAction = SavePersonnelEventInput(False)
    Application.StatusBar = "Personnel action saved: " & SavePersonnelWizardAction
End Function
Public Sub OpenPersonnelEnrollmentAction()
    mdlEnrollmentWorkflow.OpenEnrollmentForm
End Sub

Public Sub PrepareNewPersonnelAction(ByVal eventType As String)
    Dim normalizedType As String

    normalizedType = UCase$(Trim$(eventType))
    If normalizedType <> EVENT_TYPE_TRANSFER And normalizedType <> EVENT_TYPE_EXCLUSION Then
        Err.Raise vbObjectError + 659, "mdlPersonnelEvents", "Only TRANSFER and EXCLUSION can be prepared by the personnel action form."
    End If

    ResetPersonnelEventInput
    SetInputValue ThisWorkbook.Worksheets(SHEET_EVENT_INPUT), "event_type", normalizedType
End Sub

Public Sub SaveCurrentPersonnelAction()
    Dim eventID As String

    eventID = SavePersonnelEventInput(False)
    Application.StatusBar = "Personnel action saved: " & eventID
End Sub
Public Sub ExportSavedPersonnelEventOrder()
    Dim ws As Worksheet
    Dim eventID As String
    Dim outputPath As String

    EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    eventID = SafeText(GetInputValue(ws, "saved_event_id"))
    If eventID = "" Then
        Application.StatusBar = "Save the personnel action before exporting its order."
        Exit Sub
    End If
    If UCase$(SafeText(GetInputValue(ws, "event_type"))) = EVENT_TYPE_ENROLLMENT Then
        Application.StatusBar = "Use the enrollment master to export an enrollment order."
        Exit Sub
    End If

    outputPath = mdlPersonnelEventOrderExport.ExportPersonnelEventOrder(eventID)
    Application.StatusBar = "Personnel order exported: " & outputPath
End Sub

Public Sub OpenHistoryForPersonnelAction()
    Dim wsInput As Worksheet
    Dim wsHistory As Worksheet
    Dim employeeID As String

    EnsurePersonnelEventInfrastructure
    Set wsInput = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    employeeID = SafeText(GetInputValue(wsInput, "employee_id"))
    If employeeID = "" Then
        Application.StatusBar = "Save or select a personnel action before opening history."
        Exit Sub
    End If

    mdlPersonnelHistory.OpenPersonnelHistory
    Set wsHistory = ThisWorkbook.Worksheets("PersonnelHistory")
    wsHistory.Cells(4, 2).Value = employeeID
    mdlPersonnelHistory.RefreshPersonnelHistory
End Sub
Public Sub ResetPersonnelEventInput()
    Dim ws As Worksheet
    Dim lastRow As Long

    EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 4 Then ws.Range(ws.Cells(4, 2), ws.Cells(lastRow, 2)).ClearContents
    SetInputValue ws, "event_type", EVENT_TYPE_TRANSFER
    ws.Activate
End Sub

Public Sub LoadCurrentStateToPersonnelEventInput()
    Dim ws As Worksheet
    Dim employeeID As String
    Dim stateData As Object

    EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    employeeID = SafeText(GetInputValue(ws, "employee_id"))
    If employeeID = "" Then
        MsgBox PELabel("personnel.event.input.employee_required", "Enter EmployeeID before loading current state."), vbExclamation
        Exit Sub
    End If

    Set stateData = GetCurrentPersonnelState(employeeID)
    If stateData.Count = 0 Then
        MsgBox PELabel("personnel.event.input.employee_missing", "Current state was not found for this EmployeeID."), vbExclamation
        Exit Sub
    End If

    PopulateCurrentStateInput ws, stateData
    MsgBox PELabel("personnel.event.input.loaded", "Current state was loaded into the new-state fields."), vbInformation
End Sub

Public Sub PreparePersonnelEventCorrection(ByVal employeeID As String, ByVal correctsEventID As String)
    Dim ws As Worksheet
    Dim stateData As Object

    EnsurePersonnelEventInfrastructure
    employeeID = SafeText(employeeID)
    correctsEventID = SafeText(correctsEventID)
    If employeeID = "" Or correctsEventID = "" Then Err.Raise vbObjectError + 657, "mdlPersonnelEvents", "EmployeeID and EventID are required to prepare a correction."
    ValidateCorrectionTarget correctsEventID, employeeID

    Set stateData = GetCurrentPersonnelState(employeeID)
    If stateData.Count = 0 Then Err.Raise vbObjectError + 658, "mdlPersonnelEvents", "Current state was not found for EmployeeID: " & employeeID

    Set ws = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    SetInputValue ws, "event_type", EVENT_TYPE_TRANSFER
    SetInputValue ws, "employee_id", employeeID
    SetInputValue ws, "event_date", ""
    SetInputValue ws, "effective_date", ""
    SetInputValue ws, "order_reference", ""
    SetInputValue ws, "basis_text", ""
    SetInputValue ws, "comment", ""
    SetInputValue ws, "saved_event_id", ""
    SetInputValue ws, "corrects_event_id", correctsEventID
    PopulateCurrentStateInput ws, stateData
    ws.Activate
End Sub

Public Function SavePersonnelEventInput(Optional ByVal showMessage As Boolean = True) As String
    Dim ws As Worksheet
    Dim eventData As Object
    Dim beforeState As Object
    Dim afterState As Object
    Dim existingState As Object
    Dim allowanceRules As Object
    Dim paymentAssignments As Collection
    Dim eventType As String
    Dim employeeID As String

    EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    employeeID = SafeText(GetInputValue(ws, "employee_id"))
    eventType = UCase$(SafeText(GetInputValue(ws, "event_type")))
    If eventType = "" Then eventType = EVENT_TYPE_TRANSFER

    Set eventData = CreateObject("Scripting.Dictionary")
    eventData("event_type") = eventType
    eventData("event_date") = GetInputValue(ws, "event_date")
    eventData("effective_date") = GetInputValue(ws, "effective_date")
    eventData("order_reference") = GetInputValue(ws, "order_reference")
    eventData("basis_text") = GetInputValue(ws, "basis_text")
    eventData("comment") = GetInputValue(ws, "comment")
    eventData("corrects_event_id") = GetInputValue(ws, "corrects_event_id")
    eventData("handover_date") = GetInputValue(ws, "handover_date")
    eventData("acceptance_date") = GetInputValue(ws, "acceptance_date")
    eventData("duty_start_date") = GetInputValue(ws, "duty_start_date")
    eventData("destination_unit") = GetInputValue(ws, "destination_unit")
    eventData("destination_location") = GetInputValue(ws, "destination_location")
    eventData("material_assistance_status") = GetInputValue(ws, "material_assistance_status")
    eventData("main_leave_status") = GetInputValue(ws, "main_leave_status")
    eventData("additional_leave_status") = GetInputValue(ws, "additional_leave_status")

    If eventType = EVENT_TYPE_ENROLLMENT Then
        If employeeID = "" Then employeeID = BuildIdentifier("EMP")
        Set beforeState = CreateObject("Scripting.Dictionary")
        beforeState("employee_id") = employeeID
        beforeState("is_active") = "NO"
        Set afterState = BuildAfterStateFromInput(beforeState, ws)
        If SafeText(ValueOf(afterState, "fio")) = "" Then Err.Raise vbObjectError + 654, "mdlPersonnelEvents", "New FIO is required for ENROLLMENT."
        If SafeText(ValueOf(afterState, "personal_number")) = "" Then Err.Raise vbObjectError + 655, "mdlPersonnelEvents", "New personal number is required for ENROLLMENT."
        Set existingState = GetCurrentPersonnelState(employeeID)
        If existingState.Count > 0 Then Err.Raise vbObjectError + 656, "mdlPersonnelEvents", "EmployeeID already has current state: " & employeeID
        afterState("employee_id") = employeeID
        afterState("state_date") = eventData("effective_date")
        afterState("is_active") = "YES"
        Set allowanceRules = BuildAllowanceRulesFromInput(ws, eventData("effective_date"))
        Set paymentAssignments = mdlPersonnelAllowanceRules.EvaluatePersonnelAllowances(afterState, allowanceRules)
        eventData("employee_id") = employeeID
        SavePersonnelEventInput = SavePersonnelEvent(eventData, beforeState, afterState, paymentAssignments)
        SetInputValue ws, "employee_id", employeeID
    Else
        Set beforeState = GetCurrentPersonnelState(employeeID)
        If beforeState.Count = 0 Then Err.Raise vbObjectError + 652, "mdlPersonnelEvents", "Current state was not found for EmployeeID: " & employeeID
        eventData("employee_id") = employeeID
    End If

    If eventType = EVENT_TYPE_EXCLUSION Then
        Set afterState = CloneDictionary(beforeState)
        afterState("state_date") = eventData("effective_date")
        afterState("is_active") = "NO"
        SavePersonnelEventInput = SavePersonnelEvent(eventData, beforeState, afterState)
    ElseIf eventType = EVENT_TYPE_TRANSFER Then
        Set afterState = BuildAfterStateFromInput(beforeState, ws)
        afterState("state_date") = eventData("effective_date")
        Set allowanceRules = BuildAllowanceRulesFromInput(ws, eventData("effective_date"))
        Set paymentAssignments = mdlPersonnelAllowanceRules.EvaluatePersonnelAllowances(afterState, allowanceRules)
        SavePersonnelEventInput = SavePersonnelEvent(eventData, beforeState, afterState, paymentAssignments)
    ElseIf eventType <> EVENT_TYPE_ENROLLMENT Then
        Err.Raise vbObjectError + 653, "mdlPersonnelEvents", "Only ENROLLMENT, TRANSFER and EXCLUSION are supported by this input sheet."
    End If

    SetInputValue ws, "saved_event_id", SavePersonnelEventInput
    If showMessage Then MsgBox PELabel("personnel.event.input.saved", "Personnel event was saved: ") & SavePersonnelEventInput, vbInformation
End Function

Public Function SaveTransferEvent(ByVal employeeID As String, ByVal eventDate As Variant, ByVal effectiveDate As Variant, ByVal beforeState As Object, ByVal afterState As Object, ByVal orderReference As String, ByVal basisText As String, Optional ByVal paymentAssignments As Collection = Nothing) As String
    Dim eventData As Object
    Set eventData = CreateObject("Scripting.Dictionary")
    eventData("employee_id") = employeeID
    eventData("event_type") = EVENT_TYPE_TRANSFER
    eventData("event_date") = eventDate
    eventData("effective_date") = effectiveDate
    eventData("order_reference") = orderReference
    eventData("basis_text") = basisText
    SaveTransferEvent = SavePersonnelEvent(eventData, beforeState, afterState, paymentAssignments)
End Function

Public Function SaveExclusionEvent(ByVal employeeID As String, ByVal eventDate As Variant, ByVal effectiveDate As Variant, ByVal beforeState As Object, ByVal orderReference As String, ByVal basisText As String, Optional ByVal paymentAssignments As Collection = Nothing) As String
    Dim afterState As Object
    Dim eventData As Object

    Set eventData = CreateObject("Scripting.Dictionary")
    eventData("employee_id") = employeeID
    eventData("event_type") = EVENT_TYPE_EXCLUSION
    eventData("event_date") = eventDate
    eventData("effective_date") = effectiveDate
    eventData("order_reference") = orderReference
    eventData("basis_text") = basisText

    Set afterState = CloneDictionary(beforeState)
    afterState("state_date") = effectiveDate
    afterState("is_active") = "NO"
    SaveExclusionEvent = SavePersonnelEvent(eventData, beforeState, afterState, paymentAssignments)
End Function

Public Function SavePersonnelEvent(ByVal eventData As Object, ByVal beforeState As Object, ByVal afterState As Object, Optional ByVal paymentAssignments As Collection = Nothing, Optional ByVal documents As Collection = Nothing) As String
    Dim eventID As String
    Dim employeeID As String
    Dim eventType As String
    Dim eventDate As Variant
    Dim effectiveDate As Variant
    Dim beforeSnapshotID As String
    Dim afterSnapshotID As String
    Dim wsEvents As Worksheet
    Dim nextRow As Long

    EnsurePersonnelEventInfrastructure
    ValidatePersonnelEvent eventData, beforeState, afterState

    employeeID = RequiredValue(eventData, "employee_id")
    eventType = UCase$(RequiredValue(eventData, "event_type"))
    eventDate = RequiredDateValue(eventData, "event_date")
    effectiveDate = RequiredDateValue(eventData, "effective_date")

    If SafeText(ValueOf(eventData, "corrects_event_id")) <> "" Then ValidateCorrectionTarget SafeText(ValueOf(eventData, "corrects_event_id")), employeeID
    EnsureEmployee employeeID, beforeState, afterState
    eventID = BuildIdentifier("EVT")
    beforeSnapshotID = SaveStateSnapshot(eventID, "BEFORE", employeeID, beforeState)
    afterSnapshotID = SaveStateSnapshot(eventID, "AFTER", employeeID, afterState)

    Set wsEvents = ThisWorkbook.Worksheets(SHEET_EVENTS)
    nextRow = NextDataRow(wsEvents)
    WriteEventRow wsEvents, nextRow, eventID, employeeID, eventType, eventDate, effectiveDate, beforeSnapshotID, afterSnapshotID, eventData
    If SafeText(ValueOf(eventData, "corrects_event_id")) <> "" Then MarkCorrectedEvent SafeText(ValueOf(eventData, "corrects_event_id")), employeeID

    If eventType = EVENT_TYPE_EXCLUSION Then
        CloseActiveAssignments employeeID, eventID, effectiveDate
    End If

    SavePaymentAssignments employeeID, eventID, paymentAssignments
    SaveDocumentRecords eventID, documents
    UpdateCurrentState employeeID, afterState, eventID
    SavePersonnelEvent = eventID
End Function

Public Function EnsureEnrollmentPersonnelEvent(ByVal enrollmentRecord As Object) As String
    Dim employeeID As String
    Dim enrollmentID As String
    Dim existingEventID As String
    Dim eventData As Object
    Dim beforeState As Object
    Dim afterState As Object
    Dim eventDate As Date
    Dim effectiveDate As Date

    If enrollmentRecord Is Nothing Then Err.Raise vbObjectError + 661, "mdlPersonnelEvents", "Enrollment record is required."
    EnsurePersonnelEventInfrastructure
    enrollmentID = EnrollmentValue(enrollmentRecord, "enrollment_id")
    If enrollmentID = "" Then Err.Raise vbObjectError + 662, "mdlPersonnelEvents", "EnrollmentID is required."
    existingEventID = FindEnrollmentEventByEnrollmentID(enrollmentID)
    If existingEventID <> "" Then
        EnsureEnrollmentPersonnelEvent = existingEventID
        Exit Function
    End If
    eventDate = EnrollmentDate(enrollmentRecord, "order_date", Date)
    effectiveDate = EnrollmentDate(enrollmentRecord, "duty_start_date", EnrollmentDate(enrollmentRecord, "enroll_date", eventDate))
    If effectiveDate < eventDate Then effectiveDate = eventDate
    employeeID = BuildIdentifier("EMP")
    Set beforeState = CreateObject("Scripting.Dictionary")
    beforeState("employee_id") = employeeID
    beforeState("is_active") = "NO"
    Set afterState = CreateObject("Scripting.Dictionary")
    afterState("employee_id") = employeeID
    afterState("fio") = EnrollmentValue(enrollmentRecord, "fio")
    afterState("personal_number") = EnrollmentValue(enrollmentRecord, "personal_number")
    afterState("table_number") = EnrollmentValue(enrollmentRecord, "table_number")
    afterState("source_mode") = EnrollmentValue(enrollmentRecord, "source_mode")
    afterState("rank") = EnrollmentValue(enrollmentRecord, "rank")
    afterState("position") = EnrollmentValue(enrollmentRecord, "position")
    afterState("section") = EnrollmentValue(enrollmentRecord, "section")
    afterState("military_unit") = EnrollmentValue(enrollmentRecord, "military_unit")
    afterState("vus") = EnrollmentValue(enrollmentRecord, "vus")
    afterState("tariff_rank") = EnrollmentValue(enrollmentRecord, "tariff_rank")
    afterState("position_salary") = EnrollmentValue(enrollmentRecord, "position_salary")
    afterState("rank_salary") = EnrollmentValue(enrollmentRecord, "rank_salary")
    afterState("service_category") = EnrollmentValue(enrollmentRecord, "service_category")
    afterState("contract_kind") = EnrollmentValue(enrollmentRecord, "contract_kind")
    afterState("contract_basis") = EnrollmentValue(enrollmentRecord, "contract_basis")
    afterState("state_date") = effectiveDate
    afterState("is_active") = "YES"
    Set eventData = CreateObject("Scripting.Dictionary")
    eventData("employee_id") = employeeID
    eventData("event_type") = EVENT_TYPE_ENROLLMENT
    eventData("event_date") = eventDate
    eventData("effective_date") = effectiveDate
    eventData("order_reference") = EnrollmentValue(enrollmentRecord, "order_number")
    eventData("basis_text") = EnrollmentValue(enrollmentRecord, "basis_section1")
    eventData("comment") = "EnrollmentID: " & enrollmentID
    EnsureEnrollmentPersonnelEvent = SavePersonnelEvent(eventData, beforeState, afterState)
End Function

Private Function FindEnrollmentEventByEnrollmentID(ByVal enrollmentID As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim marker As String
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENTS)
    marker = "EnrollmentID: " & enrollmentID
    For rowNum = 2 To LastDataRow(ws)
        If UCase$(SafeText(ws.Cells(rowNum, 3).Value)) = EVENT_TYPE_ENROLLMENT Then
            If SafeText(ws.Cells(rowNum, 15).Value) = marker Then
                FindEnrollmentEventByEnrollmentID = SafeText(ws.Cells(rowNum, 1).Value)
                Exit Function
            End If
        End If
    Next rowNum
End Function

Private Function EnrollmentValue(ByVal enrollmentRecord As Object, ByVal fieldName As String) As String
    If enrollmentRecord.Exists(fieldName) Then EnrollmentValue = SafeText(enrollmentRecord(fieldName))
End Function

Private Function EnrollmentDate(ByVal enrollmentRecord As Object, ByVal fieldName As String, ByVal fallbackValue As Date) As Date
    If enrollmentRecord.Exists(fieldName) Then
        If IsDate(enrollmentRecord(fieldName)) Then
            EnrollmentDate = CDate(enrollmentRecord(fieldName))
            Exit Function
        End If
    End If
    EnrollmentDate = fallbackValue
End Function
Public Function GetCurrentPersonnelState(ByVal employeeID As String) As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim result As Object

    EnsurePersonnelEventInfrastructure
    Set result = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(SHEET_CURRENT_STATE)
    lastRow = LastDataRow(ws)

    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = employeeID Then
            result("employee_id") = employeeID
            result("rank") = ws.Cells(rowNum, 2).Value
            result("rank_effective_date") = ws.Cells(rowNum, 3).Value
            result("position") = ws.Cells(rowNum, 4).Value
            result("section") = ws.Cells(rowNum, 5).Value
            result("military_unit") = ws.Cells(rowNum, 6).Value
            result("vus") = ws.Cells(rowNum, 7).Value
            result("tariff_rank") = ws.Cells(rowNum, 8).Value
            result("position_salary") = ws.Cells(rowNum, 9).Value
            result("rank_salary") = ws.Cells(rowNum, 10).Value
            result("service_category") = ws.Cells(rowNum, 11).Value
            result("contract_kind") = ws.Cells(rowNum, 12).Value
            result("contract_basis") = ws.Cells(rowNum, 13).Value
            result("state_date") = ws.Cells(rowNum, 14).Value
            result("source_event_id") = ws.Cells(rowNum, 15).Value
            result("last_event_id") = ws.Cells(rowNum, 16).Value
            result("fizo_level") = ws.Cells(rowNum, 17).Value
            result("sport_status") = ws.Cells(rowNum, 18).Value
            result("medal_code") = ws.Cells(rowNum, 19).Value
            result("driver_c_d_ce") = ws.Cells(rowNum, 20).Value
            result("contract_430_eligible") = ws.Cells(rowNum, 21).Value
            result("medal_award_date") = ws.Cells(rowNum, 22).Value
            result("medal_award_document_reference") = ws.Cells(rowNum, 23).Value
            Exit For
        End If
    Next rowNum

    Set GetCurrentPersonnelState = result
End Function

Public Sub SetPersonnelEventStatus(ByVal eventID As String, ByVal eventStatus As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    EnsurePersonnelEventInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENTS)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = eventID Then
            ws.Cells(rowNum, 6).Value = UCase$(eventStatus)
            ws.Cells(rowNum, 13).Value = Now
            Exit Sub
        End If
    Next rowNum
    Err.Raise vbObjectError + 657, "mdlPersonnelEvents", "Personnel event was not found: " & eventID
End Sub

Private Sub MarkCorrectedEvent(ByVal correctedEventID As String, ByVal employeeID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_EVENTS)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = correctedEventID Then
            If SafeText(ws.Cells(rowNum, 2).Value) <> employeeID Then Err.Raise vbObjectError + 658, "mdlPersonnelEvents", "Corrected event belongs to another employee."
            ws.Cells(rowNum, 6).Value = EVENT_STATUS_CORRECTED
            ws.Cells(rowNum, 13).Value = Now
            Exit Sub
        End If
    Next rowNum
    Err.Raise vbObjectError + 659, "mdlPersonnelEvents", "Corrected event was not found: " & correctedEventID
End Sub

Private Sub ValidateCorrectionTarget(ByVal correctedEventID As String, ByVal employeeID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_EVENTS)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = correctedEventID Then
            If SafeText(ws.Cells(rowNum, 2).Value) <> employeeID Then Err.Raise vbObjectError + 658, "mdlPersonnelEvents", "Corrected event belongs to another employee."
            Exit Sub
        End If
    Next rowNum
    Err.Raise vbObjectError + 659, "mdlPersonnelEvents", "Corrected event was not found: " & correctedEventID
End Sub

Private Sub ValidatePersonnelEvent(ByVal eventData As Object, ByVal beforeState As Object, ByVal afterState As Object)
    Dim eventType As String
    Dim eventDate As Date
    Dim effectiveDate As Date

    If eventData Is Nothing Then Err.Raise vbObjectError + 640, "mdlPersonnelEvents", "Event data is required."
    If beforeState Is Nothing Then Err.Raise vbObjectError + 641, "mdlPersonnelEvents", "Before-state snapshot is required."
    If afterState Is Nothing Then Err.Raise vbObjectError + 642, "mdlPersonnelEvents", "After-state snapshot is required."

    eventType = UCase$(RequiredValue(eventData, "event_type"))
    If eventType <> EVENT_TYPE_ENROLLMENT And eventType <> EVENT_TYPE_TRANSFER And eventType <> EVENT_TYPE_EXCLUSION Then
        Err.Raise vbObjectError + 643, "mdlPersonnelEvents", "Unsupported personnel event type: " & eventType
    End If

    eventDate = CDate(RequiredDateValue(eventData, "event_date"))
    effectiveDate = CDate(RequiredDateValue(eventData, "effective_date"))
    If effectiveDate < eventDate Then Err.Raise vbObjectError + 644, "mdlPersonnelEvents", "Effective date cannot be earlier than event date."
    If eventType = EVENT_TYPE_TRANSFER Then
        If IsDate(ValueOf(eventData, "handover_date")) And IsDate(ValueOf(eventData, "acceptance_date")) Then If CDate(ValueOf(eventData, "handover_date")) > CDate(ValueOf(eventData, "acceptance_date")) Then Err.Raise vbObjectError + 645, "mdlPersonnelEvents", "Handover date cannot be later than acceptance date."
        If IsDate(ValueOf(eventData, "acceptance_date")) And IsDate(ValueOf(eventData, "duty_start_date")) Then If CDate(ValueOf(eventData, "duty_start_date")) < CDate(ValueOf(eventData, "acceptance_date")) Then Err.Raise vbObjectError + 646, "mdlPersonnelEvents", "Duty start date cannot be earlier than acceptance date."
    ElseIf eventType = EVENT_TYPE_EXCLUSION Then
        If IsDate(ValueOf(eventData, "handover_date")) Then If effectiveDate < CDate(ValueOf(eventData, "handover_date")) Then Err.Raise vbObjectError + 647, "mdlPersonnelEvents", "Exclusion date cannot be earlier than handover date."
    End If
End Sub

Private Sub EnsureEmployee(ByVal employeeID As String, ByVal beforeState As Object, ByVal afterState As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim sourceState As Object

    Set ws = ThisWorkbook.Worksheets(SHEET_EMPLOYEES)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = employeeID Then
            ws.Cells(rowNum, 9).Value = Now
            ws.Cells(rowNum, 10).Value = ValueOrDefault(afterState, "is_active", "YES")
            Exit Sub
        End If
    Next rowNum

    Set sourceState = afterState
    If SafeText(ValueOf(sourceState, "fio")) = "" Then Set sourceState = beforeState
    rowNum = NextDataRow(ws)
    ws.Cells(rowNum, 1).Value = employeeID
    ws.Cells(rowNum, 2).Value = ValueOf(sourceState, "fio")
    ws.Cells(rowNum, 3).Value = ValueOf(sourceState, "personal_number")
    ws.Cells(rowNum, 4).Value = ValueOf(sourceState, "table_number")
    ws.Cells(rowNum, 5).Value = ValueOf(sourceState, "source_mode")
    ws.Cells(rowNum, 6).Value = ValueOf(sourceState, "staff_link_status")
    ws.Cells(rowNum, 7).Value = ValueOf(sourceState, "staff_reference")
    ws.Cells(rowNum, 8).Value = Now
    ws.Cells(rowNum, 9).Value = Now
    ws.Cells(rowNum, 10).Value = ValueOrDefault(afterState, "is_active", "YES")
End Sub

Private Function SaveStateSnapshot(ByVal eventID As String, ByVal snapshotKind As String, ByVal employeeID As String, ByVal stateData As Object) As String
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim snapshotID As String

    snapshotID = BuildIdentifier("SNP")
    Set ws = ThisWorkbook.Worksheets(SHEET_SNAPSHOTS)
    nextRow = NextDataRow(ws)
    ws.Cells(nextRow, 1).Value = snapshotID
    ws.Cells(nextRow, 2).Value = eventID
    ws.Cells(nextRow, 3).Value = snapshotKind
    ws.Cells(nextRow, 4).Value = employeeID
    ws.Cells(nextRow, 5).Value = ValueOf(stateData, "rank")
    ws.Cells(nextRow, 6).Value = ValueOf(stateData, "position")
    ws.Cells(nextRow, 7).Value = ValueOf(stateData, "section")
    ws.Cells(nextRow, 8).Value = ValueOf(stateData, "military_unit")
    ws.Cells(nextRow, 9).Value = ValueOf(stateData, "vus")
    ws.Cells(nextRow, 10).Value = ValueOf(stateData, "tariff_rank")
    ws.Cells(nextRow, 11).Value = ValueOf(stateData, "position_salary")
    ws.Cells(nextRow, 12).Value = ValueOf(stateData, "rank_salary")
    ws.Cells(nextRow, 13).Value = ValueOf(stateData, "service_category")
    ws.Cells(nextRow, 14).Value = ValueOf(stateData, "contract_kind")
    ws.Cells(nextRow, 15).Value = ValueOf(stateData, "contract_basis")
    ws.Cells(nextRow, 16).Value = ValueOf(stateData, "state_date")
    ws.Cells(nextRow, 17).Value = SerializeDictionary(stateData)
    ws.Cells(nextRow, 18).Value = Now
    SaveStateSnapshot = snapshotID
End Function

Private Sub WriteEventRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal eventID As String, ByVal employeeID As String, ByVal eventType As String, ByVal eventDate As Variant, ByVal effectiveDate As Variant, ByVal beforeSnapshotID As String, ByVal afterSnapshotID As String, ByVal eventData As Object)
    ws.Cells(rowNum, 1).Value = eventID
    ws.Cells(rowNum, 2).Value = employeeID
    ws.Cells(rowNum, 3).Value = eventType
    ws.Cells(rowNum, 4).Value = eventDate
    ws.Cells(rowNum, 5).Value = effectiveDate
    ws.Cells(rowNum, 6).Value = ValueOrDefault(eventData, "status", EVENT_STATUS_DRAFT)
    ws.Cells(rowNum, 7).Value = beforeSnapshotID
    ws.Cells(rowNum, 8).Value = afterSnapshotID
    ws.Cells(rowNum, 9).Value = ValueOf(eventData, "order_reference")
    ws.Cells(rowNum, 10).Value = ValueOf(eventData, "basis_text")
    ws.Cells(rowNum, 11).Value = Application.UserName
    ws.Cells(rowNum, 12).Value = Now
    ws.Cells(rowNum, 13).Value = Now
    ws.Cells(rowNum, 14).Value = ValueOf(eventData, "corrects_event_id")
    ws.Cells(rowNum, 15).Value = ValueOf(eventData, "comment")
    ws.Cells(rowNum, 16).Value = ValueOf(eventData, "handover_date")
    ws.Cells(rowNum, 17).Value = ValueOf(eventData, "acceptance_date")
    ws.Cells(rowNum, 18).Value = ValueOf(eventData, "duty_start_date")
    ws.Cells(rowNum, 19).Value = ValueOf(eventData, "destination_unit")
    ws.Cells(rowNum, 20).Value = ValueOf(eventData, "destination_location")
    ws.Cells(rowNum, 21).Value = ValueOf(eventData, "material_assistance_status")
    ws.Cells(rowNum, 22).Value = ValueOf(eventData, "main_leave_status")
    ws.Cells(rowNum, 23).Value = ValueOf(eventData, "additional_leave_status")
End Sub

Private Sub UpdateCurrentState(ByVal employeeID As String, ByVal stateData As Object, ByVal eventID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_CURRENT_STATE)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = employeeID Then Exit For
    Next rowNum
    If rowNum > lastRow Then rowNum = NextDataRow(ws)

    ws.Cells(rowNum, 1).Value = employeeID
    ws.Cells(rowNum, 2).Value = ValueOf(stateData, "rank")
    ws.Cells(rowNum, 3).Value = ValueOf(stateData, "rank_effective_date")
    ws.Cells(rowNum, 4).Value = ValueOf(stateData, "position")
    ws.Cells(rowNum, 5).Value = ValueOf(stateData, "section")
    ws.Cells(rowNum, 6).Value = ValueOf(stateData, "military_unit")
    ws.Cells(rowNum, 7).Value = ValueOf(stateData, "vus")
    ws.Cells(rowNum, 8).Value = ValueOf(stateData, "tariff_rank")
    ws.Cells(rowNum, 9).Value = ValueOf(stateData, "position_salary")
    ws.Cells(rowNum, 10).Value = ValueOf(stateData, "rank_salary")
    ws.Cells(rowNum, 11).Value = ValueOf(stateData, "service_category")
    ws.Cells(rowNum, 12).Value = ValueOf(stateData, "contract_kind")
    ws.Cells(rowNum, 13).Value = ValueOf(stateData, "contract_basis")
    ws.Cells(rowNum, 14).Value = ValueOf(stateData, "state_date")
    ws.Cells(rowNum, 15).Value = eventID
    ws.Cells(rowNum, 16).Value = eventID
    ws.Cells(rowNum, 17).Value = ValueOf(stateData, "fizo_level")
    ws.Cells(rowNum, 18).Value = ValueOf(stateData, "sport_status")
    ws.Cells(rowNum, 19).Value = ValueOf(stateData, "medal_code")
    ws.Cells(rowNum, 20).Value = ValueOf(stateData, "driver_c_d_ce")
    ws.Cells(rowNum, 21).Value = ValueOf(stateData, "contract_430_eligible")
    ws.Cells(rowNum, 22).Value = ValueOf(stateData, "medal_award_date")
    ws.Cells(rowNum, 23).Value = ValueOf(stateData, "medal_award_document_reference")
End Sub

Private Sub SavePaymentAssignments(ByVal employeeID As String, ByVal eventID As String, ByVal paymentAssignments As Collection)
    Dim assignment As Object
    Dim ws As Worksheet
    Dim nextRow As Long

    If paymentAssignments Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_ASSIGNMENTS)
    For Each assignment In paymentAssignments
        nextRow = NextDataRow(ws)
        ws.Cells(nextRow, 1).Value = BuildIdentifier("PAY")
        ws.Cells(nextRow, 2).Value = employeeID
        ws.Cells(nextRow, 3).Value = eventID
        ws.Cells(nextRow, 4).Value = ValueOf(assignment, "payment_type")
        ws.Cells(nextRow, 5).Value = ValueOf(assignment, "payment_code")
        ws.Cells(nextRow, 6).Value = ValueOf(assignment, "amount_kind")
        ws.Cells(nextRow, 7).Value = ValueOf(assignment, "amount_value")
        ws.Cells(nextRow, 8).Value = ValueOf(assignment, "calculation_base")
        ws.Cells(nextRow, 9).Value = ValueOf(assignment, "start_date")
        ws.Cells(nextRow, 10).Value = ValueOf(assignment, "end_date")
        ws.Cells(nextRow, 11).Value = ValueOrDefault(assignment, "status", "ACTIVE")
        ws.Cells(nextRow, 12).Value = ValueOf(assignment, "termination_event_id")
        ws.Cells(nextRow, 13).Value = ValueOf(assignment, "act_id")
        ws.Cells(nextRow, 14).Value = ValueOf(assignment, "act_point")
        ws.Cells(nextRow, 15).Value = ValueOf(assignment, "factual_basis")
        ws.Cells(nextRow, 16).Value = ValueOf(assignment, "document_reference")
        ws.Cells(nextRow, 17).Value = ValueOf(assignment, "original_amount")
        ws.Cells(nextRow, 18).Value = ValueOf(assignment, "applied_amount")
        ws.Cells(nextRow, 19).Value = ValueOf(assignment, "cap_group")
        ws.Cells(nextRow, 20).Value = ValueOf(assignment, "explanation")
        ws.Cells(nextRow, 21).Value = Now
        ws.Cells(nextRow, 22).Value = Now
    Next assignment
End Sub

Private Sub CloseActiveAssignments(ByVal employeeID As String, ByVal terminationEventID As String, ByVal terminationDate As Variant)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_ASSIGNMENTS)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 2).Value) = employeeID And UCase$(SafeText(ws.Cells(rowNum, 11).Value)) = "ACTIVE" Then
            ws.Cells(rowNum, 10).Value = terminationDate
            ws.Cells(rowNum, 11).Value = "TERMINATED"
            ws.Cells(rowNum, 12).Value = terminationEventID
            ws.Cells(rowNum, 22).Value = Now
        End If
    Next rowNum
End Sub

Private Sub SaveDocumentRecords(ByVal eventID As String, ByVal documents As Collection)
    Dim documentData As Object
    Dim ws As Worksheet
    Dim nextRow As Long

    If documents Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_DOCUMENTS)
    For Each documentData In documents
        nextRow = NextDataRow(ws)
        ws.Cells(nextRow, 1).Value = BuildIdentifier("DOC")
        ws.Cells(nextRow, 2).Value = eventID
        ws.Cells(nextRow, 3).Value = ValueOf(documentData, "document_type")
        ws.Cells(nextRow, 4).Value = ValueOf(documentData, "document_number")
        ws.Cells(nextRow, 5).Value = ValueOf(documentData, "document_date")
        ws.Cells(nextRow, 6).Value = ValueOf(documentData, "file_path")
        ws.Cells(nextRow, 7).Value = ValueOf(documentData, "template_name")
        ws.Cells(nextRow, 8).Value = ValueOf(documentData, "template_version")
        ws.Cells(nextRow, 9).Value = ValueOf(documentData, "file_hash")
        ws.Cells(nextRow, 10).Value = ValueOrDefault(documentData, "status", "DRAFT")
        ws.Cells(nextRow, 11).Value = ValueOf(documentData, "last_error")
        ws.Cells(nextRow, 12).Value = Now
    Next documentData
End Sub

Private Sub EnsureServiceSheet(ByVal sheetName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Dim index As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    For index = LBound(headers) To UBound(headers)
        If SafeText(ws.Cells(1, index + 1).Value) = "" Then ws.Cells(1, index + 1).Value = headers(index)
        ws.Cells(1, index + 1).Font.Bold = True
    Next index
    'Service sheets remain filter-free. Excel may persist both _FilterDatabase
    'and _xlnm._FilterDatabase for a programmatically applied filter, which
    'causes a name-conflict prompt when the workbook is opened.
    ws.Rows(1).Interior.Color = RGB(217, 225, 242)
End Sub

Private Sub EnsurePersonnelEventInputSheet()
    Dim ws As Worksheet
    Dim fields As Variant
    Dim labels As Variant
    Dim index As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_EVENT_INPUT)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SHEET_EVENT_INPUT
    End If

    fields = Array("event_type", "employee_id", "event_date", "effective_date", "order_reference", "basis_text", "comment", "new_rank", "new_position", "new_section", "new_military_unit", "new_vus", "new_tariff_rank", "new_position_salary", "new_rank_salary", "new_service_category", "new_contract_kind", "new_contract_basis", "fizo_level", "sport_status", "medal_code", "driver_c_d_ce", "contract_430_eligible", "mobilized_fixed_act_id", "saved_event_id", "new_fio", "new_personal_number", "new_table_number", "new_source_mode", "corrects_event_id", "handover_date", "acceptance_date", "duty_start_date", "destination_unit", "destination_location", "material_assistance_status", "main_leave_status", "additional_leave_status", "medal_award_date", "medal_award_document_reference")
    labels = Array("Event type (ENROLLMENT, TRANSFER or EXCLUSION)", "Employee ID (optional for ENROLLMENT)", "Event date", "Effective date", "Order reference", "Basis text", "Comment", "New rank", "New position", "New section", "New military unit", "New VUS", "New tariff rank", "New position salary", "New rank salary", "New service category", "New contract kind", "New contract basis", "FIZO level (SECOND, FIRST or HIGH)", "Sport status (CLASS_1, CMS, MASTER)", "Medal code", "Driver C/D/CE (YES or NO)", "Contract 430 eligible (YES or NO)", "Legal act ID for 158000", "Saved event ID", "New FIO", "New personal number", "New table number", "New source mode", "Corrects saved Event ID", "Handover date", "Acceptance date", "Duty start date", "Destination unit", "Destination location", "Material assistance status", "Main leave status", "Additional leave status", "Medal award date", "Medal award-order reference")

    If SafeText(ws.Cells(1, 1).Value) = "" Then
        ws.Cells(1, 1).Value = PELabel("personnel.event.input.title", "Personnel event input")
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
    End If
    ws.Cells(3, 1).Value = "Field"
    ws.Cells(3, 2).Value = "Value"
    ws.Rows(3).Font.Bold = True
    ws.Rows(3).Interior.Color = RGB(217, 225, 242)

    For index = LBound(fields) To UBound(fields)
        ws.Cells(index + 4, 1).Value = fields(index)
        If SafeText(ws.Cells(index + 4, 3).Value) = "" Then ws.Cells(index + 4, 3).Value = labels(index)
    Next index
    If SafeText(GetInputValue(ws, "event_type")) = "" Then SetInputValue ws, "event_type", EVENT_TYPE_TRANSFER
    ws.Columns(1).ColumnWidth = 26
    ws.Columns(2).ColumnWidth = 32
    ws.Columns(3).ColumnWidth = 44
End Sub

Private Function BuildAfterStateFromInput(ByVal beforeState As Object, ByVal ws As Worksheet) As Object
    Dim result As Object
    Dim fieldNames As Variant
    Dim index As Long
    Dim newValue As Variant

    Set result = CloneDictionary(beforeState)
    fieldNames = Array("rank", "position", "section", "military_unit", "vus", "tariff_rank", "position_salary", "rank_salary", "service_category", "contract_kind", "contract_basis")
    For index = LBound(fieldNames) To UBound(fieldNames)
        newValue = GetInputValue(ws, "new_" & fieldNames(index))
        If SafeText(newValue) <> "" Then result(fieldNames(index)) = newValue
    Next index
    If SafeText(GetInputValue(ws, "new_fio")) <> "" Then result("fio") = GetInputValue(ws, "new_fio")
    If SafeText(GetInputValue(ws, "new_personal_number")) <> "" Then result("personal_number") = GetInputValue(ws, "new_personal_number")
    If SafeText(GetInputValue(ws, "new_table_number")) <> "" Then result("table_number") = GetInputValue(ws, "new_table_number")
    If SafeText(GetInputValue(ws, "new_source_mode")) <> "" Then
        result("source_mode") = GetInputValue(ws, "new_source_mode")
    ElseIf SafeText(ValueOf(result, "source_mode")) = "" Then
        result("source_mode") = "MANUAL"
    End If
    If SafeText(ValueOf(result, "staff_link_status")) = "" Then result("staff_link_status") = "MANUAL_ONLY"
    SetStateValueIfProvided result, "fizo_level", GetInputValue(ws, "fizo_level")
    SetStateValueIfProvided result, "sport_status", GetInputValue(ws, "sport_status")
    SetStateValueIfProvided result, "medal_code", GetInputValue(ws, "medal_code")
    SetStateValueIfProvided result, "medal_award_date", GetInputValue(ws, "medal_award_date")
    SetStateValueIfProvided result, "medal_award_document_reference", GetInputValue(ws, "medal_award_document_reference")
    SetStateValueIfProvided result, "driver_c_d_ce", GetInputValue(ws, "driver_c_d_ce")
    SetStateValueIfProvided result, "contract_430_eligible", GetInputValue(ws, "contract_430_eligible")
    Set BuildAfterStateFromInput = result
End Function

Private Sub PopulateCurrentStateInput(ByVal ws As Worksheet, ByVal stateData As Object)
    SetInputValue ws, "new_rank", ValueOf(stateData, "rank")
    SetInputValue ws, "new_position", ValueOf(stateData, "position")
    SetInputValue ws, "new_section", ValueOf(stateData, "section")
    SetInputValue ws, "new_military_unit", ValueOf(stateData, "military_unit")
    SetInputValue ws, "new_vus", ValueOf(stateData, "vus")
    SetInputValue ws, "new_tariff_rank", ValueOf(stateData, "tariff_rank")
    SetInputValue ws, "new_position_salary", ValueOf(stateData, "position_salary")
    SetInputValue ws, "new_rank_salary", ValueOf(stateData, "rank_salary")
    SetInputValue ws, "new_service_category", ValueOf(stateData, "service_category")
    SetInputValue ws, "new_contract_kind", ValueOf(stateData, "contract_kind")
    SetInputValue ws, "new_contract_basis", ValueOf(stateData, "contract_basis")
    SetInputValue ws, "fizo_level", ValueOf(stateData, "fizo_level")
    SetInputValue ws, "sport_status", ValueOf(stateData, "sport_status")
    SetInputValue ws, "medal_code", ValueOf(stateData, "medal_code")
    SetInputValue ws, "medal_award_date", ValueOf(stateData, "medal_award_date")
    SetInputValue ws, "medal_award_document_reference", ValueOf(stateData, "medal_award_document_reference")
    SetInputValue ws, "driver_c_d_ce", ValueOf(stateData, "driver_c_d_ce")
    SetInputValue ws, "contract_430_eligible", ValueOf(stateData, "contract_430_eligible")
End Sub

Private Sub SetStateValueIfProvided(ByVal stateData As Object, ByVal key As String, ByVal newValue As Variant)
    If SafeText(newValue) <> "" Then stateData(key) = newValue
End Sub

Private Function BuildAllowanceRulesFromInput(ByVal ws As Worksheet, ByVal effectiveDate As Variant) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("mobilized_fixed_act_id") = GetInputValue(ws, "mobilized_fixed_act_id")
    result("medal_award_date") = GetInputValue(ws, "medal_award_date")
    result("medal_award_document_reference") = GetInputValue(ws, "medal_award_document_reference")
    result("default_start_date") = effectiveDate
    Set BuildAllowanceRulesFromInput = result
End Function

Private Function FindInputFieldRow(ByVal ws As Worksheet, ByVal fieldKey As String) As Long
    Dim rowNum As Long
    For rowNum = 4 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LCase$(SafeText(ws.Cells(rowNum, 1).Value)) = LCase$(fieldKey) Then
            FindInputFieldRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Function GetInputValue(ByVal ws As Worksheet, ByVal fieldKey As String) As Variant
    Dim rowNum As Long
    rowNum = FindInputFieldRow(ws, fieldKey)
    If rowNum > 0 Then GetInputValue = ws.Cells(rowNum, 2).Value
End Function

Private Sub SetInputValue(ByVal ws As Worksheet, ByVal fieldKey As String, ByVal fieldValue As Variant)
    Dim rowNum As Long
    rowNum = FindInputFieldRow(ws, fieldKey)
    If rowNum > 0 Then ws.Cells(rowNum, 2).Value = fieldValue
End Sub

Private Function PELabel(ByVal key As String, ByVal fallback As String) As String
    On Error GoTo Fallback
    PELabel = ModuleLocalization.t(key, fallback)
    Exit Function
Fallback:
    PELabel = fallback
End Function

Private Function RequiredValue(ByVal source As Object, ByVal key As String) As String
    RequiredValue = SafeText(ValueOf(source, key))
    If RequiredValue = "" Then Err.Raise vbObjectError + 650, "mdlPersonnelEvents", "Required field is blank: " & key
End Function

Private Function RequiredDateValue(ByVal source As Object, ByVal key As String) As Variant
    RequiredDateValue = ValueOf(source, key)
    If Not IsDate(RequiredDateValue) Then Err.Raise vbObjectError + 651, "mdlPersonnelEvents", "Required date is invalid: " & key
End Function

Private Function ValueOf(ByVal source As Object, ByVal key As String) As Variant
    If source Is Nothing Then Exit Function
    If source.Exists(key) Then ValueOf = source(key)
End Function

Private Function ValueOrDefault(ByVal source As Object, ByVal key As String, ByVal fallbackValue As String) As String
    ValueOrDefault = SafeText(ValueOf(source, key))
    If ValueOrDefault = "" Then ValueOrDefault = fallbackValue
End Function

Private Function CloneDictionary(ByVal source As Object) As Object
    Dim result As Object
    Dim itemKey As Variant

    Set result = CreateObject("Scripting.Dictionary")
    If Not source Is Nothing Then
        For Each itemKey In source.Keys
            result(CStr(itemKey)) = source(itemKey)
        Next itemKey
    End If
    Set CloneDictionary = result
End Function

Private Function SerializeDictionary(ByVal source As Object) As String
    Dim itemKey As Variant
    Dim result As String

    If source Is Nothing Then Exit Function
    For Each itemKey In source.Keys
        If result <> "" Then result = result & " | "
        result = result & CStr(itemKey) & "=" & Replace$(SafeText(source(itemKey)), "|", "/")
    Next itemKey
    SerializeDictionary = result
End Function

Private Function BuildIdentifier(ByVal prefix As String) As String
    Randomize
    BuildIdentifier = prefix & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(CLng(Rnd() * 9999), "0000")
End Function

Private Function NextDataRow(ByVal ws As Worksheet) As Long
    NextDataRow = LastDataRow(ws) + 1
    If NextDataRow < 2 Then NextDataRow = 2
End Function

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LastDataRow < 1 Then LastDataRow = 1
End Function

Private Function SafeText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    SafeText = Trim$(CStr(rawValue))
End Function
