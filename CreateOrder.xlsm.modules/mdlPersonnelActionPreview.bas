Attribute VB_Name = "mdlPersonnelActionPreview"
Option Explicit

' Pure read-only preview builder for TRANSFER and EXCLUSION personnel actions.
' This module must never write workbook cells or create service sheets.

Private Const SHEET_EMPLOYEES As String = "Employees"
Private Const SHEET_CURRENT_STATE As String = "EmployeeCurrentState"
Private Const SHEET_ASSIGNMENTS As String = "PaymentAssignments"

Private Const PREVIEW_ERROR_DRAFT_REQUIRED As String = "PREVIEW_DRAFT_REQUIRED"
Private Const PREVIEW_ERROR_TYPE As String = "PREVIEW_EVENT_TYPE_INVALID"
Private Const PREVIEW_ERROR_EMPLOYEE As String = "PREVIEW_EMPLOYEE_REQUIRED"
Private Const PREVIEW_ERROR_STATE As String = "PREVIEW_STATE_NOT_FOUND"
Private Const PREVIEW_ERROR_DATE As String = "PREVIEW_DATE_INVALID"
Private Const PREVIEW_ERROR_DATE_ORDER As String = "PREVIEW_DATE_ORDER_INVALID"
Private Const PREVIEW_ERROR_SHEET As String = "PREVIEW_REGISTRY_UNAVAILABLE"
Private Const PREVIEW_ERROR_UNEXPECTED As String = "PREVIEW_UNEXPECTED_ERROR"

Public Function BuildPersonnelActionPreview(ByVal draftValues As Object) As Object
    Dim result As Object
    Dim beforeState As Object
    Dim afterState As Object
    Dim eventData As Object
    Dim paymentChanges As Collection
    Dim eventType As String
    Dim employeeID As String
    Dim errorCount As Long
    Dim changedCount As Long

    Set result = NewPreviewResult()
    On Error GoTo Failed

    If draftValues Is Nothing Then
        AddPreviewWarning result, "ERROR", PREVIEW_ERROR_DRAFT_REQUIRED, "Draft values are required."
        result("is_valid") = False
        result("can_confirm") = False
        LogPreviewError PREVIEW_ERROR_DRAFT_REQUIRED
        Set BuildPersonnelActionPreview = result
        Exit Function
    End If

    eventType = UCase$(Trim$(SafeText(DraftValue(draftValues, "event_type"))))
    employeeID = Trim$(SafeText(DraftValue(draftValues, "employee_id")))
    result("event_type") = eventType
    result("employee_id") = employeeID

    If eventType <> mdlPersonnelEvents.EVENT_TYPE_TRANSFER And eventType <> mdlPersonnelEvents.EVENT_TYPE_EXCLUSION Then
        AddPreviewWarning result, "ERROR", PREVIEW_ERROR_TYPE, "Only TRANSFER and EXCLUSION are supported."
    End If
    If employeeID = "" Then
        AddPreviewWarning result, "ERROR", PREVIEW_ERROR_EMPLOYEE, "EmployeeID is required."
    End If

    Set eventData = BuildEventData(draftValues, eventType, employeeID)
    result.Add "event_data", eventData

    If employeeID <> "" Then
        Set beforeState = ReadCurrentPersonnelStateNoWrite(employeeID)
        If beforeState.Count = 0 Then
            AddPreviewWarning result, "ERROR", PREVIEW_ERROR_STATE, "Current personnel state was not found."
        Else
            result.Add "before", beforeState
            Set afterState = BuildAfterState(beforeState, draftValues, eventData, eventType)
            result.Add "after", afterState
            AddStateDiff result, beforeState, afterState, eventData, draftValues, eventType, changedCount
            Set paymentChanges = BuildPaymentChanges(beforeState, afterState, eventData, eventType)
            Set result.Item("payment_changes") = paymentChanges
            BuildOrderProjection result, beforeState, afterState, eventData, paymentChanges
            If changedCount = 0 Then AddPreviewWarning result, "WARN", "NO_STATE_CHANGE", "The projected state is unchanged."
        End If
    End If

    ValidatePreviewDates result, eventData, eventType
    errorCount = CountWarningsBySeverity(result("warnings"), "ERROR")
    result("is_valid") = (errorCount = 0)
    result("can_confirm") = (errorCount = 0 And eventType <> "" And employeeID <> "" And result.Exists("before") And result.Exists("after"))
    UpdatePreviewCounts result, changedCount

    LogPreviewBuilt eventType, changedCount, paymentChanges, result("warnings")
    Set BuildPersonnelActionPreview = result
    Exit Function

Failed:
    If result Is Nothing Then Set result = NewPreviewResult()
    AddPreviewWarning result, "ERROR", PREVIEW_ERROR_UNEXPECTED, "Preview calculation failed."
    result("is_valid") = False
    result("can_confirm") = False
    LogPreviewError PREVIEW_ERROR_UNEXPECTED
    Set BuildPersonnelActionPreview = result
End Function

Private Function NewPreviewResult() As Object
    Dim result As Object
    Dim counts As Object
    Dim changedFields As Collection
    Dim paymentChanges As Collection
    Dim warnings As Collection

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set changedFields = New Collection
    Set paymentChanges = New Collection
    Set warnings = New Collection
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    counts.Add "changed_fields", 0
    counts.Add "payment_starts", 0
    counts.Add "payment_stops", 0
    counts.Add "payment_decisions", 0
    counts.Add "warnings", 0
    result.Add "changed_fields", changedFields
    result.Add "payment_changes", paymentChanges
    result.Add "warnings", warnings
    result.Add "counts", counts
    result.Add "is_valid", False
    result.Add "can_confirm", False
    Set NewPreviewResult = result
End Function

Private Function BuildEventData(ByVal draftValues As Object, ByVal eventType As String, ByVal employeeID As String) As Object
    Dim eventData As Object
    Dim fields As Variant
    Dim index As Long

    Set eventData = CreateObject("Scripting.Dictionary")
    eventData.CompareMode = vbTextCompare
    eventData.Add "event_type", eventType
    eventData.Add "employee_id", employeeID
    fields = Array("event_date", "effective_date", "order_reference", "basis_text", "comment", "corrects_event_id", "handover_date", "acceptance_date", "duty_start_date", "destination_unit", "destination_location", "material_assistance_status", "main_leave_status", "additional_leave_status")
    For index = LBound(fields) To UBound(fields)
        eventData.Add CStr(fields(index)), DraftValue(draftValues, CStr(fields(index)))
    Next index
    Set BuildEventData = eventData
End Function

Private Function ReadCurrentPersonnelStateNoWrite(ByVal employeeID As String) As Object
    Dim result As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employees As Worksheet

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set ws = ReadOnlyWorksheet(SHEET_CURRENT_STATE)
    lastRow = LastDataRowNoWrite(ws)
    For rowNum = 2 To lastRow
        If StrComp(SafeText(ws.Cells(rowNum, 1).Value), employeeID, vbTextCompare) = 0 Then
            result.Add "employee_id", employeeID
            result.Add "rank", ws.Cells(rowNum, 2).Value
            result.Add "rank_effective_date", ws.Cells(rowNum, 3).Value
            result.Add "position", ws.Cells(rowNum, 4).Value
            result.Add "section", ws.Cells(rowNum, 5).Value
            result.Add "military_unit", ws.Cells(rowNum, 6).Value
            result.Add "vus", ws.Cells(rowNum, 7).Value
            result.Add "tariff_rank", ws.Cells(rowNum, 8).Value
            result.Add "position_salary", ws.Cells(rowNum, 9).Value
            result.Add "rank_salary", ws.Cells(rowNum, 10).Value
            result.Add "service_category", ws.Cells(rowNum, 11).Value
            result.Add "contract_kind", ws.Cells(rowNum, 12).Value
            result.Add "contract_basis", ws.Cells(rowNum, 13).Value
            result.Add "state_date", ws.Cells(rowNum, 14).Value
            result.Add "source_event_id", ws.Cells(rowNum, 15).Value
            result.Add "last_event_id", ws.Cells(rowNum, 16).Value
            result.Add "fizo_level", ws.Cells(rowNum, 17).Value
            result.Add "sport_status", ws.Cells(rowNum, 18).Value
            result.Add "medal_code", ws.Cells(rowNum, 19).Value
            result.Add "driver_c_d_ce", ws.Cells(rowNum, 20).Value
            result.Add "contract_430_eligible", ws.Cells(rowNum, 21).Value
            result.Add "medal_award_date", ws.Cells(rowNum, 22).Value
            result.Add "medal_award_document_reference", ws.Cells(rowNum, 23).Value
            Exit For
        End If
    Next rowNum
    If result.Count = 0 Then
        Set ReadCurrentPersonnelStateNoWrite = result
        Exit Function
    End If

    Set employees = ReadOnlyWorksheet(SHEET_EMPLOYEES)
    lastRow = LastDataRowNoWrite(employees)
    For rowNum = 2 To lastRow
        If StrComp(SafeText(employees.Cells(rowNum, 1).Value), employeeID, vbTextCompare) = 0 Then
            result("fio") = employees.Cells(rowNum, 2).Value
            result("personal_number") = employees.Cells(rowNum, 3).Value
            result("table_number") = employees.Cells(rowNum, 4).Value
            result("source_mode") = employees.Cells(rowNum, 5).Value
            result("staff_link_status") = employees.Cells(rowNum, 6).Value
            result("staff_reference") = employees.Cells(rowNum, 7).Value
            result("is_active") = employees.Cells(rowNum, 10).Value
            Exit For
        End If
    Next rowNum
    Set ReadCurrentPersonnelStateNoWrite = result
End Function

Private Function BuildAfterState(ByVal beforeState As Object, ByVal draftValues As Object, ByVal eventData As Object, ByVal eventType As String) As Object
    Dim result As Object
    Dim stateFields As Variant
    Dim index As Long
    Dim stateKey As String
    Dim draftKey As String
    Dim newValue As Variant

    Set result = CloneDictionaryNoWrite(beforeState)
    If eventType = mdlPersonnelEvents.EVENT_TYPE_EXCLUSION Then
        result("state_date") = DraftValue(eventData, "effective_date")
        result("is_active") = "NO"
        Set BuildAfterState = result
        Exit Function
    End If

    stateFields = Array("rank", "position", "section", "military_unit", "vus", "tariff_rank", "position_salary", "rank_salary", "service_category", "contract_kind", "contract_basis")
    For index = LBound(stateFields) To UBound(stateFields)
        stateKey = CStr(stateFields(index))
        draftKey = "new_" & stateKey
        newValue = DraftValue(draftValues, draftKey)
        If SafeText(newValue) <> "" Then result(stateKey) = newValue
    Next index
    SetStateValueIfProvided result, "fio", DraftValue(draftValues, "new_fio")
    SetStateValueIfProvided result, "personal_number", DraftValue(draftValues, "new_personal_number")
    SetStateValueIfProvided result, "table_number", DraftValue(draftValues, "new_table_number")
    If SafeText(DraftValue(draftValues, "new_source_mode")) <> "" Then
        result("source_mode") = DraftValue(draftValues, "new_source_mode")
    ElseIf SafeText(ValueOfNoWrite(result, "source_mode")) = "" Then
        result("source_mode") = "MANUAL"
    End If
    If SafeText(ValueOfNoWrite(result, "staff_link_status")) = "" Then result("staff_link_status") = "MANUAL_ONLY"
    SetStateValueIfProvided result, "fizo_level", DraftValue(draftValues, "fizo_level")
    SetStateValueIfProvided result, "sport_status", DraftValue(draftValues, "sport_status")
    SetStateValueIfProvided result, "medal_code", DraftValue(draftValues, "medal_code")
    SetStateValueIfProvided result, "medal_award_date", DraftValue(draftValues, "medal_award_date")
    SetStateValueIfProvided result, "medal_award_document_reference", DraftValue(draftValues, "medal_award_document_reference")
    SetStateValueIfProvided result, "driver_c_d_ce", DraftValue(draftValues, "driver_c_d_ce")
    SetStateValueIfProvided result, "contract_430_eligible", DraftValue(draftValues, "contract_430_eligible")
    result("state_date") = DraftValue(eventData, "effective_date")
    Set BuildAfterState = result
End Function

Private Sub AddStateDiff(ByVal result As Object, ByVal beforeState As Object, ByVal afterState As Object, ByVal eventData As Object, ByVal draftValues As Object, ByVal eventType As String, ByRef changedCount As Long)
    Dim fields As Variant
    Dim fieldKey As Variant
    Dim beforeValue As Variant
    Dim afterValue As Variant
    Dim stateKey As String

    fields = PreviewFieldKeys(eventType)
    For Each fieldKey In fields
        If CStr(fieldKey) <> "status" Then
            stateKey = StateKeyForPreviewField(CStr(fieldKey))
            If stateKey <> "" Then
                beforeValue = PreviewBeforeValue(CStr(fieldKey), stateKey, beforeState, eventData)
                afterValue = PreviewAfterValue(CStr(fieldKey), stateKey, afterState, eventData, draftValues)
                AddFieldDiff result("changed_fields"), CStr(fieldKey), beforeValue, afterValue, changedCount
            End If
        End If
    Next fieldKey
End Sub

Private Function PreviewFieldKeys(ByVal eventType As String) As Variant
    If eventType = mdlPersonnelEvents.EVENT_TYPE_EXCLUSION Then
        PreviewFieldKeys = Array("employee_id", "event_date", "effective_date", "order_reference", "basis_text", "comment", "handover_date", "destination_unit", "destination_location", "material_assistance_status", "main_leave_status", "additional_leave_status")
    Else
        PreviewFieldKeys = Array("employee_id", "event_date", "effective_date", "order_reference", "basis_text", "comment", "new_rank", "new_position", "new_section", "new_military_unit", "new_vus", "handover_date", "acceptance_date", "duty_start_date", "destination_unit", "destination_location")
    End If
End Function

Private Function StateKeyForPreviewField(ByVal fieldKey As String) As String
    Select Case fieldKey
        Case "new_rank": StateKeyForPreviewField = "rank"
        Case "new_position": StateKeyForPreviewField = "position"
        Case "new_section": StateKeyForPreviewField = "section"
        Case "new_military_unit": StateKeyForPreviewField = "military_unit"
        Case "new_vus": StateKeyForPreviewField = "vus"
        Case "employee_id": StateKeyForPreviewField = "employee_id"
        Case Else: StateKeyForPreviewField = fieldKey
    End Select
End Function

Private Function PreviewBeforeValue(ByVal fieldKey As String, ByVal stateKey As String, ByVal beforeState As Object, ByVal eventData As Object) As Variant
    If fieldKey = "employee_id" Then
        PreviewBeforeValue = ValueOfNoWrite(beforeState, "employee_id")
    ElseIf IsStatePreviewField(fieldKey) Then
        PreviewBeforeValue = ValueOfNoWrite(beforeState, stateKey)
    Else
        PreviewBeforeValue = Empty
    End If
End Function

Private Function PreviewAfterValue(ByVal fieldKey As String, ByVal stateKey As String, ByVal afterState As Object, ByVal eventData As Object, ByVal draftValues As Object) As Variant
    If IsStatePreviewField(fieldKey) Then
        PreviewAfterValue = ValueOfNoWrite(afterState, stateKey)
    Else
        PreviewAfterValue = DraftValue(eventData, fieldKey)
    End If
End Function

Private Function IsStatePreviewField(ByVal fieldKey As String) As Boolean
    IsStatePreviewField = (fieldKey = "new_rank" Or fieldKey = "new_position" Or fieldKey = "new_section" Or fieldKey = "new_military_unit" Or fieldKey = "new_vus")
End Function

Private Sub AddFieldDiff(ByVal target As Collection, ByVal fieldKey As String, ByVal beforeValue As Variant, ByVal afterValue As Variant, ByRef changedCount As Long)
    Dim item As Object
    Dim changeKind As String

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    If ValuesEqualNoWrite(beforeValue, afterValue) Then
        changeKind = "UNCHANGED"
    ElseIf SafeText(beforeValue) = "" Then
        changeKind = "ADDED"
    ElseIf SafeText(afterValue) = "" Then
        changeKind = "REMOVED"
    Else
        changeKind = "CHANGED"
    End If
    item.Add "key", fieldKey
    item.Add "label", fieldKey
    item.Add "before", beforeValue
    item.Add "after", afterValue
    item.Add "change_kind", changeKind
    target.Add item
    If changeKind <> "UNCHANGED" Then changedCount = changedCount + 1
End Sub

Private Function BuildPaymentChanges(ByVal beforeState As Object, ByVal afterState As Object, ByVal eventData As Object, ByVal eventType As String) As Collection
    Dim results As New Collection
    Dim activeAssignments As Object
    Dim projected As Collection
    Dim ruleData As Object
    Dim item As Object
    Dim allowance As Object
    Dim key As Variant
    Dim statusValue As String
    Dim stage As String

    On Error GoTo FailedPayments

    stage = "active-read"
    Set activeAssignments = ReadActiveAssignmentsNoWrite(SafeText(ValueOfNoWrite(beforeState, "employee_id")))
    If eventType = mdlPersonnelEvents.EVENT_TYPE_EXCLUSION Then
        For Each key In activeAssignments.Keys
            Set item = activeAssignments(key)
            AddPaymentChange results, item, "STOP", item
        Next key
        Set BuildPaymentChanges = results
        Exit Function
    End If

    stage = "rule-data"
    Set ruleData = BuildRuleData(eventData, afterState)
    stage = "evaluate"
    Set projected = mdlPersonnelAllowanceRules.EvaluatePersonnelAllowances(afterState, ruleData)
    stage = "projected-loop"
    For Each allowance In projected
        statusValue = UCase$(SafeText(ValueOfNoWrite(allowance, "status")))
        key = UCase$(SafeText(ValueOfNoWrite(allowance, "payment_code")))
        If statusValue = mdlPersonnelAllowanceRules.ALLOWANCE_STATUS_NOT_APPLICABLE Then
            If activeAssignments.Exists(key) Then AddPaymentChange results, activeAssignments(key), "STOP", allowance
        ElseIf statusValue = mdlPersonnelAllowanceRules.ALLOWANCE_STATUS_REQUIRES_DECISION Then
            AddPaymentChange results, ExistingOrEmpty(activeAssignments, key), "REQUIRES_DECISION", allowance
        ElseIf statusValue = mdlPersonnelAllowanceRules.ALLOWANCE_STATUS_ACTIVE Then
            If activeAssignments.Exists(key) Then
                AddPaymentChange results, activeAssignments(key), "CONTINUE", allowance
            Else
                AddPaymentChange results, ExistingOrEmpty(activeAssignments, key), "START", allowance
            End If
        End If
        If activeAssignments.Exists(key) Then activeAssignments.Remove key
    Next allowance
    For Each key In activeAssignments.Keys
        Set item = activeAssignments(key)
        AddPaymentChange results, item, "STOP", item
    Next key
    Set BuildPaymentChanges = results
    Exit Function

FailedPayments:
    Err.Raise Err.Number, "BuildPaymentChanges:" & stage, Err.Description
End Function

Private Function BuildRuleData(ByVal eventData As Object, ByVal afterState As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result.Add "mobilized_fixed_act_id", ""
    result.Add "medal_award_date", ValueOfNoWrite(afterState, "medal_award_date")
    result.Add "medal_award_document_reference", ValueOfNoWrite(afterState, "medal_award_document_reference")
    result.Add "default_start_date", ValueOfNoWrite(eventData, "effective_date")
    Set BuildRuleData = result
End Function

Private Function ReadActiveAssignmentsNoWrite(ByVal employeeID As String) As Object
    Dim result As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim paymentCode As String
    Dim assignment As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If employeeID = "" Then
        Set ReadActiveAssignmentsNoWrite = result
        Exit Function
    End If
    Set ws = ReadOnlyWorksheet(SHEET_ASSIGNMENTS)
    lastRow = LastDataRowNoWrite(ws)
    For rowNum = 2 To lastRow
        If StrComp(SafeText(ws.Cells(rowNum, 2).Value), employeeID, vbTextCompare) = 0 Then
            If UCase$(SafeText(ws.Cells(rowNum, 11).Value)) = "ACTIVE" Then
                paymentCode = UCase$(SafeText(ws.Cells(rowNum, 5).Value))
                If paymentCode <> "" Then
                    Set assignment = CreateObject("Scripting.Dictionary")
                    assignment.CompareMode = vbTextCompare
                    assignment.Add "payment_type", ws.Cells(rowNum, 4).Value
                    assignment.Add "payment_code", ws.Cells(rowNum, 5).Value
                    assignment.Add "amount_kind", ws.Cells(rowNum, 6).Value
                    assignment.Add "amount_value", ws.Cells(rowNum, 7).Value
                    assignment.Add "cap_group", ws.Cells(rowNum, 19).Value
                    assignment.Add "act_id", ws.Cells(rowNum, 13).Value
                    assignment.Add "act_point", ws.Cells(rowNum, 14).Value
                    assignment.Add "explanation", ws.Cells(rowNum, 20).Value
                    If result.Exists(paymentCode) Then
                        Set result.Item(paymentCode) = assignment
                    Else
                        result.Add paymentCode, assignment
                    End If
                End If
            End If
        End If
    Next rowNum
    Set ReadActiveAssignmentsNoWrite = result
End Function

Private Function ExistingOrEmpty(ByVal assignments As Object, ByVal paymentCode As String) As Object
    Dim result As Object
    If assignments.Exists(paymentCode) Then
        Set ExistingOrEmpty = assignments(paymentCode)
        Exit Function
    End If
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result.Add "payment_code", paymentCode
    Set ExistingOrEmpty = result
End Function

Private Sub AddPaymentChange(ByVal target As Collection, ByVal beforeItem As Object, ByVal changeKind As String, ByVal afterItem As Object)
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item.Add "payment_code", FirstText(afterItem, beforeItem, "payment_code")
    item.Add "change_kind", changeKind
    item.Add "amount_kind", FirstValue(afterItem, beforeItem, "amount_kind")
    item.Add "amount_value", FirstValue(afterItem, beforeItem, "amount_value")
    item.Add "cap_group", FirstValue(afterItem, beforeItem, "cap_group")
    item.Add "act_id", FirstValue(afterItem, beforeItem, "act_id")
    item.Add "act_point", FirstValue(afterItem, beforeItem, "act_point")
    item.Add "explanation", FirstValue(afterItem, beforeItem, "explanation")
    target.Add item
End Sub

Private Sub BuildOrderProjection(ByVal result As Object, ByVal beforeState As Object, ByVal afterState As Object, ByVal eventData As Object, ByVal paymentChanges As Collection)
    Dim projection As Object
    Dim textModel As Object
    Set projection = CreateObject("Scripting.Dictionary")
    projection.CompareMode = vbTextCompare
    projection.Add "event_type", ValueOfNoWrite(eventData, "event_type")
    projection.Add "event_date", ValueOfNoWrite(eventData, "event_date")
    projection.Add "effective_date", ValueOfNoWrite(eventData, "effective_date")
    projection.Add "order_reference", ValueOfNoWrite(eventData, "order_reference")
    projection.Add "basis_text", ValueOfNoWrite(eventData, "basis_text")
    projection.Add "comment", ValueOfNoWrite(eventData, "comment")
    projection.Add "handover_date", ValueOfNoWrite(eventData, "handover_date")
    projection.Add "acceptance_date", ValueOfNoWrite(eventData, "acceptance_date")
    projection.Add "duty_start_date", ValueOfNoWrite(eventData, "duty_start_date")
    projection.Add "destination_unit", ValueOfNoWrite(eventData, "destination_unit")
    projection.Add "destination_location", ValueOfNoWrite(eventData, "destination_location")
    projection.Add "before_state", beforeState
    projection.Add "after_state", afterState
    projection.Add "payment_changes", paymentChanges
    Set textModel = mdlPersonnelOrderText.BuildPersonnelOrderTextModel(eventData, beforeState, afterState)
    projection.Add "text_model", textModel
    result.Add "order_projection", projection
End Sub

Private Sub ValidatePreviewDates(ByVal result As Object, ByVal eventData As Object, ByVal eventType As String)
    Dim eventDate As Variant
    Dim effectiveDate As Variant
    Dim handoverDate As Variant
    Dim acceptanceDate As Variant
    Dim dutyStartDate As Variant

    eventDate = ValueOfNoWrite(eventData, "event_date")
    effectiveDate = ValueOfNoWrite(eventData, "effective_date")
    If Not IsDate(eventDate) Then AddPreviewWarning result, "ERROR", PREVIEW_ERROR_DATE, "Event date is invalid."
    If Not IsDate(effectiveDate) Then AddPreviewWarning result, "ERROR", PREVIEW_ERROR_DATE, "Effective date is invalid."
    If IsDate(eventDate) And IsDate(effectiveDate) Then
        If DateValue(CDate(effectiveDate)) < DateValue(CDate(eventDate)) Then AddPreviewWarning result, "ERROR", PREVIEW_ERROR_DATE_ORDER, "Effective date cannot be earlier than event date."
    End If
    handoverDate = ValueOfNoWrite(eventData, "handover_date")
    acceptanceDate = ValueOfNoWrite(eventData, "acceptance_date")
    dutyStartDate = ValueOfNoWrite(eventData, "duty_start_date")
    If eventType = mdlPersonnelEvents.EVENT_TYPE_TRANSFER Then
        If IsDate(handoverDate) And IsDate(acceptanceDate) Then
            If DateValue(CDate(handoverDate)) > DateValue(CDate(acceptanceDate)) Then AddPreviewWarning result, "ERROR", "PREVIEW_HANDOVER_ORDER_INVALID", "Handover date cannot be later than acceptance date."
        End If
        If IsDate(acceptanceDate) And IsDate(dutyStartDate) Then
            If DateValue(CDate(dutyStartDate)) < DateValue(CDate(acceptanceDate)) Then AddPreviewWarning result, "ERROR", "PREVIEW_DUTY_START_ORDER_INVALID", "Duty start date cannot be earlier than acceptance date."
        End If
    ElseIf eventType = mdlPersonnelEvents.EVENT_TYPE_EXCLUSION Then
        If IsDate(handoverDate) And IsDate(effectiveDate) Then
            If DateValue(CDate(effectiveDate)) < DateValue(CDate(handoverDate)) Then AddPreviewWarning result, "ERROR", "PREVIEW_EXCLUSION_DATE_ORDER_INVALID", "Exclusion date cannot be earlier than handover date."
        End If
    End If
End Sub

Private Sub AddPreviewWarning(ByVal result As Object, ByVal severity As String, ByVal code As String, ByVal detailText As String)
    Dim warning As Object
    Set warning = CreateObject("Scripting.Dictionary")
    warning.CompareMode = vbTextCompare
    warning.Add "severity", severity
    warning.Add "code", code
    warning.Add "label", code
    warning.Add "detail", detailText
    result("warnings").Add warning
End Sub

Private Sub UpdatePreviewCounts(ByVal result As Object, ByVal changedCount As Long)
    Dim warning As Object
    Dim payment As Object
    Dim counts As Object
    Set counts = result("counts")
    counts("changed_fields") = changedCount
    For Each payment In result("payment_changes")
        Select Case UCase$(SafeText(ValueOfNoWrite(payment, "change_kind")))
            Case "START": counts("payment_starts") = CLng(counts("payment_starts")) + 1
            Case "STOP": counts("payment_stops") = CLng(counts("payment_stops")) + 1
            Case "REQUIRES_DECISION": counts("payment_decisions") = CLng(counts("payment_decisions")) + 1
        End Select
    Next payment
    For Each warning In result("warnings")
        counts("warnings") = CLng(counts("warnings")) + 1
    Next warning
End Sub

Private Function CountWarningsBySeverity(ByVal warnings As Collection, ByVal severity As String) As Long
    Dim warning As Object
    For Each warning In warnings
        If UCase$(SafeText(ValueOfNoWrite(warning, "severity"))) = UCase$(severity) Then CountWarningsBySeverity = CountWarningsBySeverity + 1
    Next warning
End Function

Private Function CloneDictionaryNoWrite(ByVal source As Object) As Object
    Dim result As Object
    Dim key As Variant
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If Not source Is Nothing Then
        For Each key In source.Keys
            result(CStr(key)) = source(key)
        Next key
    End If
    Set CloneDictionaryNoWrite = result
End Function

Private Function ReadOnlyWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set ReadOnlyWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ReadOnlyWorksheet Is Nothing Then Err.Raise vbObjectError + 670, "mdlPersonnelActionPreview", PREVIEW_ERROR_SHEET
End Function

Private Function LastDataRowNoWrite(ByVal ws As Worksheet) As Long
    LastDataRowNoWrite = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LastDataRowNoWrite < 1 Then LastDataRowNoWrite = 1
End Function

Private Function DraftValue(ByVal source As Object, ByVal key As String) As Variant
    If source Is Nothing Then Exit Function
    On Error GoTo Missing
    If source.Exists(key) Then DraftValue = source(key)
Missing:
End Function

Private Function ValueOfNoWrite(ByVal source As Object, ByVal key As String) As Variant
    If source Is Nothing Then Exit Function
    On Error GoTo Missing
    If source.Exists(key) Then ValueOfNoWrite = source(key)
Missing:
End Function

Private Function SafeText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    SafeText = Trim$(CStr(rawValue))
End Function

Private Function ValuesEqualNoWrite(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsDate(leftValue) And IsDate(rightValue) Then
        ValuesEqualNoWrite = (DateValue(CDate(leftValue)) = DateValue(CDate(rightValue)))
    ElseIf IsNumeric(leftValue) And IsNumeric(rightValue) And SafeText(leftValue) <> "" And SafeText(rightValue) <> "" Then
        ValuesEqualNoWrite = (CDbl(leftValue) = CDbl(rightValue))
    Else
        ValuesEqualNoWrite = (UCase$(NormalizeSpaces(SafeText(leftValue))) = UCase$(NormalizeSpaces(SafeText(rightValue))))
    End If
End Function

Private Function NormalizeSpaces(ByVal valueText As String) As String
    valueText = Trim$(valueText)
    Do While InStr(1, valueText, "  ", vbBinaryCompare) > 0
        valueText = Replace$(valueText, "  ", " ")
    Loop
    NormalizeSpaces = valueText
End Function

Private Sub SetStateValueIfProvided(ByVal stateData As Object, ByVal key As String, ByVal newValue As Variant)
    If SafeText(newValue) <> "" Then stateData(key) = newValue
End Sub

Private Function FirstValue(ByVal primary As Object, ByVal fallback As Object, ByVal key As String) As Variant
    If Not primary Is Nothing Then
        If primary.Exists(key) Then
            If SafeText(primary(key)) <> "" Then
                FirstValue = primary(key)
                Exit Function
            End If
        End If
    End If
    FirstValue = ValueOfNoWrite(fallback, key)
End Function

Private Function FirstText(ByVal primary As Object, ByVal fallback As Object, ByVal key As String) As String
    FirstText = SafeText(FirstValue(primary, fallback, key))
End Function

Private Sub LogPreviewBuilt(ByVal eventType As String, ByVal changedCount As Long, ByVal paymentChanges As Collection, ByVal warnings As Collection)
    Debug.Print "[PERSONNEL-PREVIEW] DEBUG preview-built; event_type=" & eventType & "; changed_fields=" & CStr(changedCount) & "; payments=" & CStr(paymentChanges.Count) & "; warnings=" & CStr(warnings.Count)
End Sub

Private Sub LogPreviewError(ByVal errorCode As String)
    Debug.Print "[PERSONNEL-PREVIEW] ERROR preview-failed; code=" & errorCode
End Sub
