Attribute VB_Name = "mdlGroupedPersonnelOrderExport"
Option Explicit

' P5 grouped personnel-order layer. Payments are deliberately never merged
' between employees. Events are grouped only by event category and retain the
' source-table order.

Private Const EVENTS_SHEET As String = "PersonnelEvents"
Private Const EMPLOYEES_SHEET As String = "Employees"
Private Const ASSIGNMENTS_SHEET As String = "PaymentAssignments"
Private Const DOCUMENTS_SHEET As String = "DocumentRegistry"
Private Const LINKS_SHEET As String = "DocumentEventLinks"
Private Const GROUPED_DOCUMENT_TYPE As String = "PERSONNEL_GROUPED_ORDER"

Public Function BuildGroupedPersonnelOrderReport(Optional ByVal selectedEventIDs As String = "") As String
    Dim events As Collection
    Dim errors As Collection
    Dim warnings As Collection
    Dim paragraphCount As Long
    Dim report As String
    Dim eventData As Object
    Dim rowIndex As Long

    Set errors = New Collection
    Set warnings = New Collection
    ReadGroupedEvents events, selectedEventIDs, errors
    If events.Count = 0 And errors.Count = 0 Then errors.Add "ERROR|field=event_selection|message=No saved personnel events were selected."
    ValidateGroupedEvents events, errors, warnings

    If errors.Count > 0 Then
        report = "INVALID|error_count=" & CStr(errors.Count) & "|warning_count=" & CStr(warnings.Count) & vbCrLf
        AppendCollectionLines report, errors
        AppendCollectionLines report, warnings
        BuildGroupedPersonnelOrderReport = report
        Exit Function
    End If

    paragraphCount = GetGroupedParagraphCount(events)
    report = "OK|event_count=" & CStr(events.Count) & "|paragraph_count=" & CStr(paragraphCount) & "|warning_count=" & CStr(warnings.Count) & vbCrLf
    For rowIndex = 1 To events.Count
        Set eventData = events(rowIndex)
        If rowIndex = 1 Then
            AppendLine report, "PARAGRAPH|no=" & CStr(eventData("paragraph_no")) & "|event_type=" & ReportText(eventData("event_type"))
        ElseIf CLng(eventData("paragraph_no")) <> CLng(events(rowIndex - 1)("paragraph_no")) Then
            AppendLine report, "PARAGRAPH|no=" & CStr(eventData("paragraph_no")) & "|event_type=" & ReportText(eventData("event_type"))
        End If
        AppendLine report, "ITEM|paragraph=" & CStr(eventData("paragraph_no")) & "|item=" & CStr(eventData("item_no")) & _
            "|event_id=" & ReportText(eventData("event_id")) & "|employee_id=" & ReportText(eventData("employee_id")) & _
            "|fio=" & ReportText(GetEmployeeField(eventData("employee_id"), "FIO"))
        AppendPaymentReportLines report, CStr(eventData("event_id"))
    Next rowIndex
    AppendCollectionLines report, warnings
    BuildGroupedPersonnelOrderReport = report
End Function

Public Function ValidateGroupedPersonnelOrder(Optional ByVal selectedEventIDs As String = "") As String
    Dim report As String
    report = BuildGroupedPersonnelOrderReport(selectedEventIDs)
    If Left$(report, 8) = "INVALID|" Then
        ValidateGroupedPersonnelOrder = report
    Else
        ValidateGroupedPersonnelOrder = Replace$(Split(report, vbCrLf)(0), "OK|", "VALID|")
    End If
End Function

Public Function ExportGroupedPersonnelOrder(Optional ByVal selectedEventIDs As String = "") As String
    Dim events As Collection
    Dim errors As Collection
    Dim warnings As Collection
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim outputPath As String
    Dim documentID As String
    Dim primaryEventID As String
    Dim paragraphNo As Long
    Dim itemNo As Long
    Dim eventData As Object
    Dim firstEvent As Object
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    Set errors = New Collection
    Set warnings = New Collection
    ReadGroupedEvents events, selectedEventIDs, errors
    If events.Count = 0 And errors.Count = 0 Then errors.Add "No saved personnel events were selected."
    ValidateGroupedEvents events, errors, warnings
    If errors.Count > 0 Then Err.Raise vbObjectError + 880, "mdlGroupedPersonnelOrderExport", JoinCollection(errors, vbCrLf)

    Set firstEvent = events(1)
    primaryEventID = CStr(firstEvent("event_id"))
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add
    WriteGroupedHeader wordDoc, firstEvent

    paragraphNo = 0
    itemNo = 0
    For Each eventData In events
        If CLng(eventData("paragraph_no")) <> paragraphNo Then
            paragraphNo = CLng(eventData("paragraph_no"))
            itemNo = 0
            AppendDocParagraph wordDoc, "§" & CStr(paragraphNo) & " — " & EventTypeLabel(CStr(eventData("event_type"))), True
        End If
        itemNo = itemNo + 1
        WriteGroupedEventItem wordDoc, eventData, itemNo
    Next eventData
    WriteGroupedSignature wordDoc
    ApplyGroupedFormatting wordDoc

    outputPath = BuildGroupedOutputPath()
    wordDoc.SaveAs2 outputPath, 16
    documentID = BuildDocumentID()
    RegisterGroupedDocument documentID, primaryEventID, outputPath
    RegisterGroupedEventLinks documentID, events
    MarkGroupedEventsExported events
    ExportGroupedPersonnelOrder = outputPath

CleanExit:
    On Error Resume Next
    If failureNumber <> 0 And Len(outputPath) > 0 Then
        If Len(Dir$(outputPath)) > 0 Then Kill outputPath
    End If
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "mdlGroupedPersonnelOrderExport", failureDescription
    Exit Function

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume CleanExit
End Function

Public Sub ExportGroupedPersonnelOrderPrompt()
    Dim selectedEventIDs As String
    Dim outputPath As String
    On Error GoTo Failed
    selectedEventIDs = InputBox(t("personnel.grouped.prompt.ids", "Введите EventID через запятую или оставьте пустым для всех сохранённых записей."), _
        t("personnel.grouped.prompt.title", "Единый приказ"))
    outputPath = ExportGroupedPersonnelOrder(selectedEventIDs)
    MsgBox t("personnel.grouped.message.exported", "Единый приказ сформирован:") & vbCrLf & outputPath, vbInformation
    Exit Sub
Failed:
    MsgBox t("personnel.grouped.error.export", "Единый приказ не сформирован:") & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub ReadGroupedEvents(ByRef events As Collection, ByVal selectedEventIDs As String, ByVal errors As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim eventID As String
    Dim selected As Object
    Dim eventData As Object
    Dim paragraphByType As Object
    Dim firstItemByParagraph As Object
    Dim eventType As String
    Dim statusValue As String
    Dim paragraphNo As Long

    Set events = New Collection
    Set selected = ParseSelection(selectedEventIDs)
    Set paragraphByType = CreateObject("Scripting.Dictionary")
    paragraphByType.CompareMode = vbTextCompare
    Set firstItemByParagraph = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(EVENTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        eventID = CellText(ws, rowNum, "EventID")
        If eventID <> "" Then
            statusValue = UCase$(CellText(ws, rowNum, "Status"))
            If statusValue <> "CANCELLED" And SelectionIncludes(selected, eventID) Then
                eventType = UCase$(CellText(ws, rowNum, "EventType"))
                Set eventData = CreateObject("Scripting.Dictionary")
                eventData.CompareMode = vbTextCompare
                eventData("row_index") = rowNum
                eventData("event_id") = eventID
                eventData("employee_id") = CellText(ws, rowNum, "EmployeeID")
                eventData("event_type") = eventType
                eventData("event_date") = CellText(ws, rowNum, "EventDate")
                eventData("effective_date") = CellText(ws, rowNum, "EffectiveDate")
                eventData("order_reference") = CellText(ws, rowNum, "OrderReference")
                eventData("basis_text") = CellText(ws, rowNum, "BasisText")
                eventData("before_snapshot_id") = CellText(ws, rowNum, "BeforeSnapshotID")
                eventData("after_snapshot_id") = CellText(ws, rowNum, "AfterSnapshotID")
                If Not paragraphByType.Exists(eventType) Then
                    paragraphNo = paragraphByType.Count + 1
                    paragraphByType.Add eventType, paragraphNo
                    firstItemByParagraph.Add CStr(paragraphNo), 0
                End If
                paragraphNo = CLng(paragraphByType(eventType))
                firstItemByParagraph(CStr(paragraphNo)) = CLng(firstItemByParagraph(CStr(paragraphNo))) + 1
                eventData("paragraph_no") = paragraphNo
                eventData("item_no") = CLng(firstItemByParagraph(CStr(paragraphNo)))
                events.Add eventData
            End If
        End If
    Next rowNum
End Sub

Private Sub ValidateGroupedEvents(ByVal events As Collection, ByVal errors As Collection, ByVal warnings As Collection)
    Dim eventData As Object
    Dim fio As String
    If events Is Nothing Then Exit Sub
    For Each eventData In events
        If CStr(eventData("event_id")) = "" Then AddEventError errors, eventData, "event_id", "EventID is required."
        If CStr(eventData("employee_id")) = "" Then AddEventError errors, eventData, "employee_id", "EmployeeID is required."
        If Not IsSupportedEventType(CStr(eventData("event_type"))) Then AddEventError errors, eventData, "event_type", "Unsupported event type."
        If CStr(eventData("event_date")) = "" Then AddEventError errors, eventData, "event_date", "EventDate is required."
        If CStr(eventData("effective_date")) = "" Then AddEventError errors, eventData, "effective_date", "EffectiveDate is required."
        If CStr(eventData("event_date")) <> "" And Not IsDate(eventData("event_date")) Then AddEventError errors, eventData, "event_date", "EventDate is not a valid date."
        If CStr(eventData("effective_date")) <> "" And Not IsDate(eventData("effective_date")) Then AddEventError errors, eventData, "effective_date", "EffectiveDate is not a valid date."
        If IsDate(eventData("event_date")) And IsDate(eventData("effective_date")) Then
            If CDate(eventData("effective_date")) < CDate(eventData("event_date")) Then AddEventError errors, eventData, "effective_date", "EffectiveDate cannot be earlier than EventDate."
        End If
        fio = GetEmployeeField(CStr(eventData("employee_id")), "FIO")
        If fio = "" Then AddEventError errors, eventData, "fio", "Employee FIO was not found."
        ValidateEventPayments eventData, errors, warnings
    Next eventData
End Sub

Private Sub ValidateEventPayments(ByVal eventData As Object, ByVal errors As Collection, ByVal warnings As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim eventID As String
    Dim paymentCode As String
    Dim basisText As String
    Dim actID As String
    Dim amountKind As String
    Dim amountValue As String
    Dim totalPercent As Double
    Dim driverFound As Boolean
    Dim statusValue As String

    Set ws = ThisWorkbook.Worksheets(ASSIGNMENTS_SHEET)
    lastRow = LastDataRow(ws)
    eventID = CStr(eventData("event_id"))
    For rowNum = 2 To lastRow
        If CellText(ws, rowNum, "EventID") = eventID Then
            statusValue = UCase$(CellText(ws, rowNum, "Status"))
            If statusValue <> "CANCELLED" Then
                paymentCode = CellText(ws, rowNum, "PaymentCode")
                basisText = CellText(ws, rowNum, "FactualBasis")
                actID = CellText(ws, rowNum, "ActID")
                amountKind = UCase$(CellText(ws, rowNum, "AmountKind"))
                amountValue = CellText(ws, rowNum, "AmountValue")
                If paymentCode = "" Then AddPaymentError errors, eventData, rowNum, "payment_code", "PaymentCode is required for a selected payment."
                If paymentCode <> "" And PaymentRuleRequiresDecision(paymentCode) Then AddPaymentError errors, eventData, rowNum, "payment_code", "The payment rule requires an explicit legal decision before export."
                If basisText = "" Then AddPaymentError errors, eventData, rowNum, "factual_basis", "Не указано обязательное основание выплаты: " & paymentCode & ". Заполните поле или снимите флажок выплаты."
                If actID = "" Then AddPaymentError errors, eventData, rowNum, "act_id", "ActID is required for a selected payment."
                If actID <> "" And Not LegalActExists(actID) Then AddPaymentError errors, eventData, rowNum, "act_id", "The referenced legal act was not found."
                If amountKind = "" Then AddPaymentError errors, eventData, rowNum, "amount_kind", "AmountKind is required for a selected payment."
                If amountValue = "" Then AddPaymentError errors, eventData, rowNum, "amount_value", "AmountValue is required for a selected payment."
                If amountKind = "PERCENT" And amountValue <> "" And Not IsNumeric(amountValue) Then AddPaymentError errors, eventData, rowNum, "amount_value", "AmountValue must be numeric for a percentage payment."
                If amountKind = "PERCENT" And IsNumeric(amountValue) Then totalPercent = totalPercent + CDbl(amountValue)
                If InStr(1, UCase$(paymentCode), "DRIVER", vbTextCompare) > 0 Then driverFound = True
            End If
        End If
    Next rowNum
    If totalPercent > 100 Then
        If driverFound Then
            warnings.Add "WARNING|event_id=" & ReportText(eventID) & "|message=Общий процент выше 100% за счёт отдельной водительской надбавки; проверьте основание."
        Else
            warnings.Add "WARNING|event_id=" & ReportText(eventID) & "|message=Общий процент выше 100%; проверьте основания выплат."
        End If
    End If
End Sub

Private Function PaymentRuleRequiresDecision(ByVal paymentCode As String) As Boolean
    Dim ws As Worksheet
    Dim rowNum As Long
    On Error GoTo Missing
    Set ws = ThisWorkbook.Worksheets("PaymentRules")
    For rowNum = 2 To LastDataRow(ws)
        If StrComp(CellText(ws, rowNum, "PaymentCode"), paymentCode, vbTextCompare) = 0 Then
            If StrComp(UCase$(CellText(ws, rowNum, "RuleStatus")), "REQUIRES_DECISION", vbTextCompare) = 0 Then
                PaymentRuleRequiresDecision = True
                Exit Function
            End If
        End If
    Next rowNum
    Exit Function
Missing:
    PaymentRuleRequiresDecision = False
End Function

Private Function LegalActExists(ByVal actID As String) As Boolean
    Dim ws As Worksheet
    Dim rowNum As Long
    On Error GoTo Missing
    Set ws = ThisWorkbook.Worksheets("LegalActs")
    For rowNum = 2 To LastDataRow(ws)
        If StrComp(CellText(ws, rowNum, "ActID"), actID, vbTextCompare) = 0 Then
            LegalActExists = True
            Exit Function
        End If
    Next rowNum
    Exit Function
Missing:
    LegalActExists = False
End Function

Private Sub AppendPaymentReportLines(ByRef report As String, ByVal eventID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim statusValue As String
    Set ws = ThisWorkbook.Worksheets(ASSIGNMENTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If CellText(ws, rowNum, "EventID") = eventID Then
            statusValue = UCase$(CellText(ws, rowNum, "Status"))
            If statusValue <> "CANCELLED" Then
                AppendLine report, "PAYMENT|event_id=" & ReportText(eventID) & "|assignment_id=" & ReportText(CellText(ws, rowNum, "AssignmentID")) & _
                    "|payment_code=" & ReportText(CellText(ws, rowNum, "PaymentCode")) & "|amount_kind=" & ReportText(CellText(ws, rowNum, "AmountKind")) & _
                    "|amount=" & ReportText(CellText(ws, rowNum, "AmountValue")) & "|basis=" & ReportText(CellText(ws, rowNum, "FactualBasis"))
            End If
        End If
    Next rowNum
End Sub

Private Sub WriteGroupedHeader(ByVal wordDoc As Object, ByVal firstEvent As Object)
    Dim unitNumber As String
    Dim cityName As String
    Dim headerText As String
    Dim lineItem As Variant
    Dim headerLines As Variant
    unitNumber = GetEnrollmentSettingSafe("enrollment.unit_number", "81510")
    cityName = GetEnrollmentSettingSafe("enrollment.city", "город")
    headerText = GetEnrollmentSettingSafe("enrollment.header_text", "ПРОЕКТ ПРИКАЗА|командира воинской части {unit}")
    headerLines = Split(Replace$(headerText, "{unit}", unitNumber), "|")
    For Each lineItem In headerLines
        If Trim$(CStr(lineItem)) <> "" Then AppendDocParagraph wordDoc, Trim$(CStr(lineItem)), True, True
    Next lineItem
    AppendDocParagraph wordDoc, "от " & CStr(firstEvent("event_date")) & " г. № " & CStr(firstEvent("order_reference")), False, True
    AppendDocParagraph wordDoc, "г. " & cityName, False, True
    AppendDocParagraph wordDoc, "", False
End Sub

Private Sub WriteGroupedEventItem(ByVal wordDoc As Object, ByVal eventData As Object, ByVal itemNo As Long)
    Dim employeeID As String
    Dim fio As String
    Dim personalNumber As String
    Dim beforePosition As String
    Dim afterPosition As String
    Dim destination As String
    employeeID = CStr(eventData("employee_id"))
    fio = GetEmployeeField(employeeID, "FIO")
    personalNumber = GetEmployeeField(employeeID, "PersonalNumber")
    beforePosition = GetSnapshotField(CStr(eventData("before_snapshot_id")), "Position")
    afterPosition = GetSnapshotField(CStr(eventData("after_snapshot_id")), "Position")
    destination = GetSnapshotField(CStr(eventData("after_snapshot_id")), "MilitaryUnit")

    AppendDocParagraph wordDoc, CStr(itemNo) & ". " & EventTypeLabel(CStr(eventData("event_type"))) & ": " & fio & _
        IIf(personalNumber = "", "", ", " & L("personnel.grouped.personal_number", "личный номер") & " " & personalNumber), True
    AppendDocParagraph wordDoc, L("personnel.grouped.event_date", "Дата события") & ": " & CStr(eventData("event_date")) & "; " & _
        L("personnel.grouped.effective_date", "с даты") & ": " & CStr(eventData("effective_date")), False
    If beforePosition <> "" Then AppendDocParagraph wordDoc, L("personnel.grouped.before", "Прежняя должность") & ": " & beforePosition, False
    If afterPosition <> "" Then AppendDocParagraph wordDoc, L("personnel.grouped.after", "Новая должность") & ": " & afterPosition, False
    If destination <> "" Then AppendDocParagraph wordDoc, L("personnel.grouped.destination", "Место службы") & ": " & destination, False
    If CStr(eventData("order_reference")) <> "" Then AppendDocParagraph wordDoc, L("personnel.grouped.order", "Приказ-основание") & ": " & CStr(eventData("order_reference")), False
    AppendGroupedPaymentsToWord wordDoc, CStr(eventData("event_id"))
    If CStr(eventData("basis_text")) <> "" Then AppendDocParagraph wordDoc, L("personnel.grouped.basis", "ОСНОВАНИЕ") & ": " & CStr(eventData("basis_text")), False
    AppendDocParagraph wordDoc, "", False
End Sub

Private Sub AppendGroupedPaymentsToWord(ByVal wordDoc As Object, ByVal eventID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim statusValue As String
    Dim paymentCode As String
    Set ws = ThisWorkbook.Worksheets(ASSIGNMENTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If CellText(ws, rowNum, "EventID") = eventID Then
            statusValue = UCase$(CellText(ws, rowNum, "Status"))
            If statusValue <> "CANCELLED" Then
                paymentCode = CellText(ws, rowNum, "PaymentCode")
                AppendDocParagraph wordDoc, "- " & paymentCode & ": " & CellText(ws, rowNum, "AmountValue") & " " & _
                    CellText(ws, rowNum, "AmountKind") & "; " & L("personnel.grouped.payment_basis", "основание") & ": " & CellText(ws, rowNum, "FactualBasis"), False
            End If
        End If
    Next rowNum
End Sub

Private Sub WriteGroupedSignature(ByVal wordDoc As Object)
    AppendDocParagraph wordDoc, "", False
    AppendDocParagraph wordDoc, GetEnrollmentSettingSafe("enrollment.signatory_position", "Командир воинской части"), True
    AppendDocParagraph wordDoc, GetEnrollmentSettingSafe("enrollment.signatory_rank", ""), False
    AppendDocParagraph wordDoc, GetEnrollmentSettingSafe("enrollment.signatory_name", ""), False
End Sub

Private Sub ApplyGroupedFormatting(ByVal wordDoc As Object)
    With wordDoc.Content.Font
        .Name = "Times New Roman"
        .Size = 12
    End With
    wordDoc.Content.ParagraphFormat.SpaceAfter = 0
End Sub

Private Sub AppendDocParagraph(ByVal wordDoc As Object, ByVal textValue As String, Optional ByVal isBold As Boolean = False, Optional ByVal centered As Boolean = False)
    Dim paragraph As Object
    Set paragraph = wordDoc.Paragraphs.Add
    paragraph.Range.Text = textValue
    paragraph.Range.Font.Bold = isBold
    If centered Then paragraph.Range.ParagraphFormat.Alignment = 1
    paragraph.Range.InsertParagraphAfter
End Sub

Private Sub RegisterGroupedDocument(ByVal documentID As String, ByVal primaryEventID As String, ByVal filePath As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Set ws = ThisWorkbook.Worksheets(DOCUMENTS_SHEET)
    rowNum = LastDataRow(ws) + 1
    If rowNum < 2 Then rowNum = 2
    ws.Cells(rowNum, HeaderColumn(ws, "DocumentID")).Value = documentID
    ws.Cells(rowNum, HeaderColumn(ws, "EventID")).Value = primaryEventID
    ws.Cells(rowNum, HeaderColumn(ws, "DocumentType")).Value = GROUPED_DOCUMENT_TYPE
    ws.Cells(rowNum, HeaderColumn(ws, "FilePath")).Value = filePath
    ws.Cells(rowNum, HeaderColumn(ws, "TemplateName")).Value = "ОбразецПриказа_распознан.pdf"
    ws.Cells(rowNum, HeaderColumn(ws, "Status")).Value = "EXPORTED"
    ws.Cells(rowNum, HeaderColumn(ws, "CreatedAt")).Value = Now
End Sub

Private Sub RegisterGroupedEventLinks(ByVal documentID As String, ByVal events As Collection)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim eventData As Object
    EnsureLinksSheet
    Set ws = ThisWorkbook.Worksheets(LINKS_SHEET)
    For Each eventData In events
        rowNum = LastDataRow(ws) + 1
        If rowNum < 2 Then rowNum = 2
        ws.Cells(rowNum, HeaderColumn(ws, "DocumentID")).Value = documentID
        ws.Cells(rowNum, HeaderColumn(ws, "EventID")).Value = eventData("event_id")
        ws.Cells(rowNum, HeaderColumn(ws, "ParagraphNo")).Value = eventData("paragraph_no")
        ws.Cells(rowNum, HeaderColumn(ws, "Role")).Value = "PARAGRAPH_ITEM"
        ws.Cells(rowNum, HeaderColumn(ws, "CreatedAt")).Value = Now
    Next eventData
End Sub

Private Sub EnsureLinksSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LINKS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LINKS_SHEET
        ws.Range("A1:E1").Value = Array("DocumentID", "EventID", "ParagraphNo", "Role", "CreatedAt")
    End If
End Sub

Private Sub MarkGroupedEventsExported(ByVal events As Collection)
    Dim eventData As Object
    For Each eventData In events
        mdlPersonnelEvents.SetPersonnelEventStatus CStr(eventData("event_id")), mdlPersonnelEvents.EVENT_STATUS_EXPORTED
    Next eventData
End Sub

Private Function BuildGroupedOutputPath() As String
    Dim outputFolder As String
    outputFolder = ThisWorkbook.Path & "\PersonnelOrders"
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder
    BuildGroupedOutputPath = outputFolder & "\GroupedPersonnel_" & Format$(Now, "yyyymmdd-hhnnss") & ".docx"
End Function

Private Function BuildDocumentID() As String
    BuildDocumentID = "DOC-GRP-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(Int((Timer * 100) Mod 100))
End Function

Private Function EventTypeLabel(ByVal eventType As String) As String
    Select Case UCase$(eventType)
        Case "ENROLLMENT": EventTypeLabel = L("personnel.grouped.enrollment", "Зачисление")
        Case "TRANSFER": EventTypeLabel = L("personnel.grouped.transfer", "Перемещение")
        Case "EXCLUSION": EventTypeLabel = L("personnel.grouped.exclusion", "Исключение из списков")
        Case Else: EventTypeLabel = eventType
    End Select
End Function

Private Function IsSupportedEventType(ByVal eventType As String) As Boolean
    IsSupportedEventType = (UCase$(eventType) = "ENROLLMENT" Or UCase$(eventType) = "TRANSFER" Or UCase$(eventType) = "EXCLUSION")
End Function

Private Function GetGroupedParagraphCount(ByVal events As Collection) As Long
    Dim eventData As Object
    For Each eventData In events
        If CLng(eventData("paragraph_no")) > GetGroupedParagraphCount Then GetGroupedParagraphCount = CLng(eventData("paragraph_no"))
    Next eventData
End Function

Private Function ParseSelection(ByVal selectedEventIDs As String) As Object
    Dim result As Object
    Dim token As Variant
    Dim normalized As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    normalized = Replace$(Replace$(selectedEventIDs, ",", "|"), ";", "|")
    If Trim$(normalized) = "" Then
        result("*") = True
    Else
        For Each token In Split(normalized, "|")
            If Trim$(CStr(token)) <> "" Then result(Trim$(CStr(token))) = True
        Next token
    End If
    Set ParseSelection = result
End Function

Private Function SelectionIncludes(ByVal selected As Object, ByVal eventID As String) As Boolean
    SelectionIncludes = selected.Exists("*") Or selected.Exists(eventID)
End Function

Private Function GetEmployeeField(ByVal employeeID As String, ByVal fieldName As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Set ws = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    For rowNum = 2 To LastDataRow(ws)
        If CellText(ws, rowNum, "EmployeeID") = employeeID Then
            GetEmployeeField = CellText(ws, rowNum, fieldName)
            Exit Function
        End If
    Next rowNum
End Function

Private Function GetSnapshotField(ByVal snapshotID As String, ByVal fieldName As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long
    If snapshotID = "" Then Exit Function
    Set ws = ThisWorkbook.Worksheets("PersonnelStateSnapshots")
    For rowNum = 2 To LastDataRow(ws)
        If CellText(ws, rowNum, "SnapshotID") = snapshotID Then
            GetSnapshotField = CellText(ws, rowNum, fieldName)
            Exit Function
        End If
    Next rowNum
End Function

Private Function CellText(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal fieldName As String) As String
    Dim columnNo As Long
    columnNo = HeaderColumn(ws, fieldName)
    If columnNo > 0 Then CellText = Trim$(CStr(ws.Cells(rowNum, columnNo).Value))
End Function

Private Function HeaderColumn(ByVal ws As Worksheet, ByVal fieldName As String) As Long
    Dim columnNo As Long
    Dim lastColumn As Long
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For columnNo = 1 To lastColumn
        If StrComp(Trim$(CStr(ws.Cells(1, columnNo).Value)), fieldName, vbTextCompare) = 0 Then
            HeaderColumn = columnNo
            Exit Function
        End If
    Next columnNo
End Function

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then LastDataRow = 1 Else LastDataRow = found.Row
End Function

Private Sub AddEventError(ByVal errors As Collection, ByVal eventData As Object, ByVal fieldName As String, ByVal messageText As String)
    errors.Add "ERROR|event_id=" & ReportText(eventData("event_id")) & "|employee_id=" & ReportText(eventData("employee_id")) & _
        "|field=" & fieldName & "|message=" & ReportText(messageText)
End Sub

Private Sub AddPaymentError(ByVal errors As Collection, ByVal eventData As Object, ByVal rowNum As Long, ByVal fieldName As String, ByVal messageText As String)
    errors.Add "ERROR|event_id=" & ReportText(eventData("event_id")) & "|employee_id=" & ReportText(eventData("employee_id")) & _
        "|assignment_row=" & CStr(rowNum) & "|field=" & fieldName & "|message=" & ReportText(messageText)
End Sub

Private Function ReportText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    ReportText = Replace$(Replace$(Replace$(Trim$(CStr(rawValue)), "|", "/"), vbCr, " "), vbLf, " ")
End Function

Private Sub AppendLine(ByRef buffer As String, ByVal lineText As String)
    buffer = buffer & lineText & vbCrLf
End Sub

Private Sub AppendCollectionLines(ByRef buffer As String, ByVal values As Collection)
    Dim value As Variant
    For Each value In values
        AppendLine buffer, CStr(value)
    Next value
End Sub

Private Function JoinCollection(ByVal values As Collection, ByVal separator As String) As String
    Dim result As String
    Dim value As Variant
    For Each value In values
        If result <> "" Then result = result & separator
        result = result & CStr(value)
    Next value
    JoinCollection = result
End Function

Private Function L(ByVal key As String, ByVal fallback As String) As String
    On Error Resume Next
    L = t(key, fallback)
    If L = "" Then L = fallback
    On Error GoTo 0
End Function

Private Function GetEnrollmentSettingSafe(ByVal key As String, ByVal fallback As String) As String
    On Error Resume Next
    GetEnrollmentSettingSafe = mdlEnrollmentWorkflow.GetEnrollmentSetting(key, fallback)
    If GetEnrollmentSettingSafe = "" Then GetEnrollmentSettingSafe = fallback
    On Error GoTo 0
End Function
