Attribute VB_Name = "mdlPersonnelEventOrderExport"
Option Explicit

Private Const EVENTS_SHEET As String = "PersonnelEvents"
Private Const SNAPSHOTS_SHEET As String = "PersonnelStateSnapshots"
Private Const EMPLOYEES_SHEET As String = "Employees"
Private Const ASSIGNMENTS_SHEET As String = "PaymentAssignments"
Private Const DOCUMENTS_SHEET As String = "DocumentRegistry"
Private Const LEGAL_ACT_MO_727 As String = "MO-727-20191206"
Private Const LEGAL_ACT_MO_430 As String = "MO-430-DSP-20190731"
Private Const LEGAL_ACT_UP_788 As String = "UP-788-20221102"

Public Function ExportPersonnelEventOrder(ByVal eventID As String) As String
    Dim eventData As Object
    Dim beforeState As Object
    Dim afterState As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim outputPath As String

    Set eventData = GetEvent(eventID)
    If eventData.Count = 0 Then Err.Raise vbObjectError + 710, "mdlPersonnelEventOrderExport", "Personnel event was not found: " & eventID
    Set beforeState = GetSnapshot(eventData("before_snapshot_id"))
    Set afterState = GetSnapshot(eventData("after_snapshot_id"))

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add

    WriteHeader wordDoc
    If eventData("event_type") = "TRANSFER" Then
        WriteTransferOrder wordDoc, eventData, beforeState, afterState
    ElseIf eventData("event_type") = "EXCLUSION" Then
        WriteExclusionOrder wordDoc, eventData, beforeState, afterState
    Else
        Err.Raise vbObjectError + 711, "mdlPersonnelEventOrderExport", "Word export is supported only for TRANSFER and EXCLUSION."
    End If

    outputPath = BuildOutputPath(eventID)
    wordDoc.SaveAs2 outputPath, 16
    RegisterDocument eventID, outputPath
    mdlPersonnelEvents.SetPersonnelEventStatus eventID, mdlPersonnelEvents.EVENT_STATUS_EXPORTED
    ExportPersonnelEventOrder = outputPath

SafeExit:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    If Err.Number <> 0 Then Err.Raise Err.Number, "mdlPersonnelEventOrderExport", Err.Description
End Function

Public Sub ExportPersonnelEventOrderPrompt()
    Dim eventID As String
    Dim outputPath As String

    On Error GoTo ErrorHandler
    eventID = Trim$(InputBox(Txt("personnel.word.prompt.event_id", "Enter saved personnel event ID:"), Txt("personnel.word.prompt.title", "Personnel order export")))
    If eventID = "" Then Exit Sub
    outputPath = ExportPersonnelEventOrder(eventID)
    MsgBox Txt("personnel.word.message.exported", "Personnel order was exported to:") & vbCrLf & outputPath, vbInformation
    Exit Sub
ErrorHandler:
    MsgBox Txt("personnel.word.error.export", "Personnel order export failed:") & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub WriteHeader(ByVal wordDoc As Object)
    AppendCenteredParagraph wordDoc, "ПРОЕКТ ПРИКАЗА", True
    AppendCenteredParagraph wordDoc, "КОМАНДИРА ВОЙСКОВОЙ ЧАСТИ", True
    AppendCenteredParagraph wordDoc, "(по строевой части)", False
    AppendCenteredParagraph wordDoc, "«___» __________ 20__ г. № ____", False
    AppendParagraph wordDoc, "", False
End Sub

Private Sub WriteTransferOrder(ByVal wordDoc As Object, ByVal eventData As Object, ByVal beforeState As Object, ByVal afterState As Object)
    Dim fio As String
    Dim coreText As String

    fio = GetEmployeeFio(eventData("employee_id"))
    AppendParagraph wordDoc, "§ 1", True
    coreText = SnapshotText(afterState, "rank") & " " & fio & ", " & SnapshotText(beforeState, "position") & _
        ", назначенного " & eventData("order_reference") & " на воинскую должность " & SnapshotText(afterState, "position") & _
        ", ВУС-" & SnapshotText(afterState, "vus") & ", полагать с " & FormatEventDate(EventDateOrFallback(eventData, "handover_date", "event_date")) & _
        " сдавшим дела и должность по предыдущей воинской должности, с " & FormatEventDate(EventDateOrFallback(eventData, "acceptance_date", "effective_date")) & _
        " принявшим дела и должность по новой воинской должности и с " & FormatEventDate(EventDateOrFallback(eventData, "duty_start_date", "effective_date")) & " вступившим в исполнение служебных обязанностей."
    AppendParagraph wordDoc, coreText, False
    AppendAllowances wordDoc, eventData("event_id"), False
    AppendBasis wordDoc, eventData("order_reference"), eventData("basis_text")
End Sub

Private Sub WriteExclusionOrder(ByVal wordDoc As Object, ByVal eventData As Object, ByVal beforeState As Object, ByVal afterState As Object)
    Dim fio As String
    Dim coreText As String

    fio = GetEmployeeFio(eventData("employee_id"))
    AppendParagraph wordDoc, "§ 1", True
    coreText = SnapshotText(beforeState, "rank") & " " & fio & ", " & SnapshotText(beforeState, "position") & _
        ", с " & FormatEventDate(EventDateOrFallback(eventData, "handover_date", "event_date")) & " сдавшим дела и должность и с " & _
        FormatEventDate(eventData("effective_date")) & " исключить из списков личного состава воинской части, всех видов обеспечения" & DestinationText(eventData) & "."
    AppendParagraph wordDoc, coreText, False
    AppendParagraph wordDoc, "Прекратить с даты исключения выплату ранее установленных надбавок и повышающих коэффициентов.", False
    AppendTerminatedAllowances wordDoc, eventData("event_id")
    AppendExclusionServiceDetails wordDoc, eventData
    AppendBasis wordDoc, eventData("order_reference"), eventData("basis_text")
End Sub

Private Sub AppendAllowances(ByVal wordDoc As Object, ByVal eventID As String, ByVal includeTerminated As Boolean)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim point2Total As Double

    Set ws = ThisWorkbook.Worksheets(ASSIGNMENTS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 3).Value) = eventID And UCase$(SafeText(ws.Cells(rowNum, 11).Value)) = "ACTIVE" And SafeText(ws.Cells(rowNum, 19).Value) = "SPECIAL_ACHIEVEMENTS_P2" Then
            If IsNumeric(ws.Cells(rowNum, 17).Value) Then point2Total = point2Total + CDbl(ws.Cells(rowNum, 17).Value)
        End If
    Next rowNum
    AppendAllowancesForAct wordDoc, ws, eventID, LEGAL_ACT_MO_727, False
    If HasActiveAllowanceForAct(ws, eventID, LEGAL_ACT_MO_430) Then
        AppendParagraph wordDoc, "В соответствии с Правилами выплаты ежемесячной надбавки за особые достижения в службе военнослужащим Вооруженных Сил Российской Федерации, проходящим военную службу по контракту, утвержденными приказом Министра обороны Российской Федерации от 31 июля 2019 г. № 430дсп, установить следующие надбавки:", False
        If point2Total > 100 Then AppendParagraph wordDoc, "Ежемесячная надбавка за особые достижения в службе по основаниям пункта 2 выплачивается в общей сумме не более 100 процентов оклада по воинской должности.", False
        AppendAllowanceRowsForAct wordDoc, ws, eventID, LEGAL_ACT_MO_430
    End If
    If HasActiveAllowanceForAct(ws, eventID, LEGAL_ACT_UP_788) Then
        AppendParagraph wordDoc, "В соответствии с Указом Президента Российской Федерации от 2 ноября 2022 г. № 788 «О ежемесячной социальной выплате гражданам Российской Федерации, призванным на военную службу по мобилизации в Вооруженные Силы Российской Федерации» установить ежемесячную социальную выплату:", False
        AppendAllowanceRowsForAct wordDoc, ws, eventID, LEGAL_ACT_UP_788
    End If
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 3).Value) = eventID And UCase$(SafeText(ws.Cells(rowNum, 11).Value)) = "ACTIVE" Then
            If SafeText(ws.Cells(rowNum, 13).Value) = "" Then AppendAllowanceLine wordDoc, ws, rowNum
        End If
    Next rowNum
End Sub

Private Sub AppendAllowancesForAct(ByVal wordDoc As Object, ByVal ws As Worksheet, ByVal eventID As String, ByVal actID As String, ByVal includeCap As Boolean)
    If Not HasActiveAllowanceForAct(ws, eventID, actID) Then Exit Sub
    If actID = LEGAL_ACT_MO_727 Then
        AppendParagraph wordDoc, "В соответствии с Порядком обеспечения денежным довольствием военнослужащих Вооруженных Сил Российской Федерации и предоставления им и членам их семей отдельных выплат, определенным приказом Министра обороны Российской Федерации от 6 декабря 2019 г. № 727, установить следующие надбавки и повышающие коэффициенты:", False
    End If
    AppendAllowanceRowsForAct wordDoc, ws, eventID, actID
End Sub

Private Function HasActiveAllowanceForAct(ByVal ws As Worksheet, ByVal eventID As String, ByVal actID As String) As Boolean
    Dim rowNum As Long
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If SafeText(ws.Cells(rowNum, 3).Value) = eventID And UCase$(SafeText(ws.Cells(rowNum, 11).Value)) = "ACTIVE" And SafeText(ws.Cells(rowNum, 13).Value) = actID Then
            HasActiveAllowanceForAct = True
            Exit Function
        End If
    Next rowNum
End Function

Private Sub AppendAllowanceRowsForAct(ByVal wordDoc As Object, ByVal ws As Worksheet, ByVal eventID As String, ByVal actID As String)
    Dim rowNum As Long
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If SafeText(ws.Cells(rowNum, 3).Value) = eventID And UCase$(SafeText(ws.Cells(rowNum, 11).Value)) = "ACTIVE" And SafeText(ws.Cells(rowNum, 13).Value) = actID Then AppendAllowanceLine wordDoc, ws, rowNum
    Next rowNum
End Sub

Private Sub AppendAllowanceLine(ByVal wordDoc As Object, ByVal ws As Worksheet, ByVal rowNum As Long)
    Dim amountText As String
    amountText = FormatAllowanceAmount(ws.Cells(rowNum, 6).Value, ws.Cells(rowNum, 18).Value)
    If SafeText(ws.Cells(rowNum, 5).Value) = "MOBILIZED_FIXED_158000" Then
        AppendParagraph wordDoc, "ежемесячную социальную выплату в размере " & amountText & " рублей.", False
    Else
        AppendParagraph wordDoc, "- " & DisplayPaymentName(SafeText(ws.Cells(rowNum, 5).Value), SafeText(ws.Cells(rowNum, 4).Value)) & ": " & amountText & ".", False
    End If
End Sub

Private Function DisplayPaymentName(ByVal paymentCode As String, ByVal fallbackName As String) As String
    Select Case paymentCode
        Case "FIZO_FIRST", "FIZO_HIGH": DisplayPaymentName = "надбавка за особые достижения в службе по физической подготовке"
        Case "TARIFF_1_4": DisplayPaymentName = "надбавка за особые достижения в службе за должность с 1-4 тарифным разрядом"
        Case "MOBILIZATION_OR_SVO_CONTRACT": DisplayPaymentName = "надбавка за особые достижения в службе по пункту 3.4 Правил"
        Case "VUS_310100_310101": DisplayPaymentName = "надбавка за особые достижения в службе по ВУС"
        Case "DRIVER_C_D_CE": DisplayPaymentName = "надбавка за особые достижения в службе водителю"
        Case Else: DisplayPaymentName = fallbackName
    End Select
End Function

Private Sub AppendTerminatedAllowances(ByVal wordDoc As Object, ByVal eventID As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(ASSIGNMENTS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 12).Value) = eventID Then
            AppendParagraph wordDoc, "- " & SafeText(ws.Cells(rowNum, 4).Value) & ".", False
        End If
    Next rowNum
End Sub

Private Sub AppendBasis(ByVal wordDoc As Object, ByVal orderReference As String, ByVal basisText As String)
    Dim textValue As String
    textValue = "ОСНОВАНИЕ: " & orderReference
    If Trim$(basisText) <> "" Then textValue = textValue & "; " & basisText
    AppendParagraph wordDoc, textValue, True
End Sub

Private Sub AppendParagraph(ByVal wordDoc As Object, ByVal textValue As String, ByVal isBold As Boolean)
    Dim paragraphRange As Object
    Set paragraphRange = wordDoc.Content
    paragraphRange.Collapse 0
    paragraphRange.InsertAfter textValue & vbCrLf
    paragraphRange.ParagraphFormat.Alignment = 0
    paragraphRange.Font.Bold = isBold
End Sub

Private Sub AppendCenteredParagraph(ByVal wordDoc As Object, ByVal textValue As String, ByVal isBold As Boolean)
    Dim paragraphRange As Object
    Set paragraphRange = wordDoc.Content
    paragraphRange.Collapse 0
    paragraphRange.InsertAfter textValue & vbCrLf
    paragraphRange.ParagraphFormat.Alignment = 1
    paragraphRange.Font.Bold = isBold
End Sub

Private Function GetEvent(ByVal eventID As String) As Object
    Dim result As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set result = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(EVENTS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = eventID Then
            result("event_id") = eventID
            result("employee_id") = SafeText(ws.Cells(rowNum, 2).Value)
            result("event_type") = SafeText(ws.Cells(rowNum, 3).Value)
            result("event_date") = ws.Cells(rowNum, 4).Value
            result("effective_date") = ws.Cells(rowNum, 5).Value
            result("before_snapshot_id") = SafeText(ws.Cells(rowNum, 7).Value)
            result("after_snapshot_id") = SafeText(ws.Cells(rowNum, 8).Value)
            result("order_reference") = SafeText(ws.Cells(rowNum, 9).Value)
            result("basis_text") = SafeText(ws.Cells(rowNum, 10).Value)
            result("handover_date") = ws.Cells(rowNum, 16).Value
            result("acceptance_date") = ws.Cells(rowNum, 17).Value
            result("duty_start_date") = ws.Cells(rowNum, 18).Value
            result("destination_unit") = SafeText(ws.Cells(rowNum, 19).Value)
            result("destination_location") = SafeText(ws.Cells(rowNum, 20).Value)
            result("material_assistance_status") = SafeText(ws.Cells(rowNum, 21).Value)
            result("main_leave_status") = SafeText(ws.Cells(rowNum, 22).Value)
            result("additional_leave_status") = SafeText(ws.Cells(rowNum, 23).Value)
            Exit For
        End If
    Next rowNum
    Set GetEvent = result
End Function

Private Function EventDateOrFallback(ByVal eventData As Object, ByVal key As String, ByVal fallbackKey As String) As Variant
    If eventData.Exists(key) Then
        If IsDate(eventData(key)) Then
            EventDateOrFallback = eventData(key)
            Exit Function
        End If
    End If
    EventDateOrFallback = eventData(fallbackKey)
End Function

Private Function DestinationText(ByVal eventData As Object) As String
    If SafeText(eventData("destination_unit")) <> "" Then DestinationText = ", полагать убывшим к новому месту службы в " & SafeText(eventData("destination_unit"))
    If SafeText(eventData("destination_location")) <> "" Then DestinationText = DestinationText & ", " & SafeText(eventData("destination_location"))
End Function

Private Sub AppendExclusionServiceDetails(ByVal wordDoc As Object, ByVal eventData As Object)
    If SafeText(eventData("material_assistance_status")) <> "" Then AppendParagraph wordDoc, "Материальная помощь за текущий год: " & SafeText(eventData("material_assistance_status")) & ".", False
    If SafeText(eventData("main_leave_status")) <> "" Then AppendParagraph wordDoc, "Основной отпуск за текущий год: " & SafeText(eventData("main_leave_status")) & ".", False
    If SafeText(eventData("additional_leave_status")) <> "" Then AppendParagraph wordDoc, "Дополнительный отпуск за текущий год: " & SafeText(eventData("additional_leave_status")) & ".", False
End Sub

Private Function GetSnapshot(ByVal snapshotID As String) As Object
    Dim result As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set result = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(SNAPSHOTS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = snapshotID Then
            result("rank") = ws.Cells(rowNum, 5).Value
            result("position") = ws.Cells(rowNum, 6).Value
            result("section") = ws.Cells(rowNum, 7).Value
            result("military_unit") = ws.Cells(rowNum, 8).Value
            result("vus") = ws.Cells(rowNum, 9).Value
            result("tariff_rank") = ws.Cells(rowNum, 10).Value
            Exit For
        End If
    Next rowNum
    Set GetSnapshot = result
End Function

Private Function GetEmployeeFio(ByVal employeeID As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If SafeText(ws.Cells(rowNum, 1).Value) = employeeID Then
            GetEmployeeFio = SafeText(ws.Cells(rowNum, 2).Value)
            Exit Function
        End If
    Next rowNum
End Function

Private Sub RegisterDocument(ByVal eventID As String, ByVal filePath As String)
    Dim ws As Worksheet
    Dim rowNum As Long
    Set ws = ThisWorkbook.Worksheets(DOCUMENTS_SHEET)
    rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If rowNum < 2 Then rowNum = 2
    ws.Cells(rowNum, 1).Value = "DOC-" & Format$(Now, "yyyymmdd-hhnnss")
    ws.Cells(rowNum, 2).Value = eventID
    ws.Cells(rowNum, 3).Value = "PERSONNEL_ORDER"
    ws.Cells(rowNum, 6).Value = filePath
    ws.Cells(rowNum, 7).Value = "PersonnelEventOrder"
    ws.Cells(rowNum, 10).Value = "EXPORTED"
    ws.Cells(rowNum, 12).Value = Now
End Sub

Private Function BuildOutputPath(ByVal eventID As String) As String
    Dim outputFolder As String
    outputFolder = ThisWorkbook.Path & "\PersonnelOrders"
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder
    BuildOutputPath = outputFolder & "\Personnel_" & eventID & ".docx"
End Function

Private Function SnapshotText(ByVal stateData As Object, ByVal key As String) As String
    If stateData.Exists(key) Then SnapshotText = SafeText(stateData(key))
End Function

Private Function FormatEventDate(ByVal rawValue As Variant) As String
    If IsDate(rawValue) Then FormatEventDate = Format$(CDate(rawValue), "dd.mm.yyyy") Else FormatEventDate = SafeText(rawValue)
End Function

Private Function FormatAllowanceAmount(ByVal amountKind As Variant, ByVal amountValue As Variant) As String
    If UCase$(SafeText(amountKind)) = "PERCENT" Then
        FormatAllowanceAmount = SafeText(amountValue) & "%"
    Else
        FormatAllowanceAmount = SafeText(amountValue)
    End If
End Function

Private Function Txt(ByVal key As String, ByVal fallback As String) As String
    On Error GoTo Fallback
    Txt = ModuleLocalization.t(key, fallback)
    Exit Function
Fallback:
    Txt = fallback
End Function

Private Function SafeText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    SafeText = Trim$(CStr(rawValue))
End Function
