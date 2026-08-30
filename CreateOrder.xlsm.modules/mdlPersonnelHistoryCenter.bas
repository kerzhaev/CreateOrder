Attribute VB_Name = "mdlPersonnelHistoryCenter"
Option Explicit

Private Const EMPLOYEES_SHEET As String = "Employees"
Private Const CURRENT_STATE_SHEET As String = "EmployeeCurrentState"
Private Const EVENTS_SHEET As String = "PersonnelEvents"
Private Const SNAPSHOTS_SHEET As String = "PersonnelStateSnapshots"
Private Const ASSIGNMENTS_SHEET As String = "PaymentAssignments"
Private Const DOCUMENTS_SHEET As String = "DocumentRegistry"
Private Const STAFF_SYNC_LOG_SHEET As String = "StaffStateSyncLog"

Public Function ResolvePersonnelHistoryEmployeeID(ByVal query As String) As String
    ResolvePersonnelHistoryEmployeeID = FindEmployeeID(Trim$(query))
End Function

Public Function BuildPersonnelHistoryCenterReport(ByVal query As String, Optional ByVal selectedEventID As String = "", Optional ByVal selectedDocumentID As String = "") As String
    Dim employeeID As String
    Dim employeeRow As Long
    Dim resultText As String
    Dim eventCount As Long
    Dim snapshotCount As Long
    Dim assignmentCount As Long
    Dim documentCount As Long
    Dim syncCount As Long

    On Error GoTo Failed
    employeeID = FindEmployeeID(Trim$(query))
    employeeRow = FindEmployeeRow(employeeID)
    If employeeRow = 0 Then Err.Raise vbObjectError + 760, "mdlPersonnelHistoryCenter", "Employee was not found."

    AppendLine resultText, "HISTORY | employee_id=" & employeeID & " | fio=" & CellText(ThisWorkbook.Worksheets(EMPLOYEES_SHEET).Cells(employeeRow, 2).Value)
    AppendLine resultText, "CARD | personal_number=" & CellText(ThisWorkbook.Worksheets(EMPLOYEES_SHEET).Cells(employeeRow, 3).Value) & " | table_number=" & CellText(ThisWorkbook.Worksheets(EMPLOYEES_SHEET).Cells(employeeRow, 4).Value) & " | active=" & CellText(ThisWorkbook.Worksheets(EMPLOYEES_SHEET).Cells(employeeRow, 10).Value)
    AppendLine resultText, "CURRENT | " & CurrentStateLine(employeeID)
    AppendLine resultText, "-- EVENTS --"
    AppendEvents resultText, employeeID, eventCount
    AppendLine resultText, "-- SNAPSHOTS --"
    AppendSnapshots resultText, employeeID, snapshotCount
    AppendLine resultText, "-- ASSIGNMENTS --"
    AppendAssignments resultText, employeeID, assignmentCount
    AppendLine resultText, "-- DOCUMENTS --"
    AppendDocuments resultText, employeeID, selectedEventID, documentCount
    AppendLine resultText, "-- STAFF SYNC --"
    AppendStaffSync resultText, employeeID, syncCount
    AppendLine resultText, "SUMMARY | events=" & CStr(eventCount) & " | snapshots=" & CStr(snapshotCount) & " | assignments=" & CStr(assignmentCount) & " | documents=" & CStr(documentCount) & " | sync=" & CStr(syncCount)
    BuildPersonnelHistoryCenterReport = resultText
    Debug.Print "INFO history-search events=" & CStr(eventCount) & " snapshots=" & CStr(snapshotCount) & " assignments=" & CStr(assignmentCount) & " documents=" & CStr(documentCount) & " sync=" & CStr(syncCount)
    Exit Function
Failed:
    Debug.Print "ERROR history-failed number=" & CStr(Err.Number)
    Err.Raise Err.Number, "mdlPersonnelHistoryCenter.BuildPersonnelHistoryCenterReport", Err.Description
End Function

Public Function GetPersonnelHistoryDocumentPath(ByVal employeeID As String, ByVal eventID As String, Optional ByVal documentID As String = "") As String
    Dim documentSheet As Worksheet
    Dim eventSheet As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim pathText As String
    Dim currentEventID As String

    On Error GoTo Failed
    If FindEmployeeRow(Trim$(employeeID)) = 0 Then Err.Raise vbObjectError + 761, "mdlPersonnelHistoryCenter", "Employee was not found."
    If Not EventBelongsToEmployee(Trim$(eventID), Trim$(employeeID)) Then Err.Raise vbObjectError + 762, "mdlPersonnelHistoryCenter", "Event does not belong to the selected employee."
    Set documentSheet = ThisWorkbook.Worksheets(DOCUMENTS_SHEET)
    lastRow = LastDataRow(documentSheet)
    For rowNum = 2 To lastRow
        currentEventID = CellText(documentSheet.Cells(rowNum, 2).Value)
        If StrComp(currentEventID, Trim$(eventID), vbTextCompare) = 0 Then
            If Len(Trim$(documentID)) = 0 Or StrComp(CellText(documentSheet.Cells(rowNum, 1).Value), Trim$(documentID), vbTextCompare) = 0 Then
                pathText = CellText(documentSheet.Cells(rowNum, 6).Value)
                If Len(pathText) = 0 Then Err.Raise vbObjectError + 763, "mdlPersonnelHistoryCenter", "Document path is empty."
                If Len(Dir$(pathText)) = 0 Then Err.Raise vbObjectError + 764, "mdlPersonnelHistoryCenter", "Registered document file is missing."
                GetPersonnelHistoryDocumentPath = pathText
                Exit Function
            End If
        End If
    Next rowNum
    Err.Raise vbObjectError + 765, "mdlPersonnelHistoryCenter", "Registered document was not found."
Failed:
    If Err.Number >= vbObjectError + 763 And Err.Number <= vbObjectError + 765 Then
        Debug.Print "WARN history-document-missing number=" & CStr(Err.Number)
    Else
        Debug.Print "ERROR history-failed number=" & CStr(Err.Number)
    End If
    Err.Raise Err.Number, "mdlPersonnelHistoryCenter.GetPersonnelHistoryDocumentPath", Err.Description
End Function

Public Sub OpenPersonnelHistoryDocument(ByVal employeeID As String, ByVal eventID As String, Optional ByVal documentID As String = "")
    Dim pathText As String
    pathText = GetPersonnelHistoryDocumentPath(employeeID, eventID, documentID)
    ThisWorkbook.FollowHyperlink Address:=pathText
    Debug.Print "INFO history-document-opened"
End Sub

Public Function RepeatPersonnelHistoryExport(ByVal employeeID As String, ByVal eventID As String) As String
    On Error GoTo Failed
    If Not EventBelongsToEmployee(Trim$(eventID), Trim$(employeeID)) Then Err.Raise vbObjectError + 766, "mdlPersonnelHistoryCenter", "Event does not belong to the selected employee."
    RepeatPersonnelHistoryExport = mdlPersonnelEventOrderExport.ExportPersonnelEventOrder(Trim$(eventID))
    Debug.Print "INFO history-export-requested"
    Exit Function
Failed:
    Debug.Print "ERROR history-failed number=" & CStr(Err.Number)
    Err.Raise Err.Number, "mdlPersonnelHistoryCenter.RepeatPersonnelHistoryExport", Err.Description
End Function

Public Sub PreparePersonnelHistoryCorrectionFromCenter(ByVal employeeID As String, ByVal eventID As String)
    On Error GoTo Failed
    If Not EventBelongsToEmployee(Trim$(eventID), Trim$(employeeID)) Then Err.Raise vbObjectError + 767, "mdlPersonnelHistoryCenter", "Event does not belong to the selected employee."
    mdlPersonnelEvents.PreparePersonnelEventCorrection Trim$(employeeID), Trim$(eventID)
    Debug.Print "INFO history-correction-prepared"
    Exit Sub
Failed:
    Debug.Print "ERROR history-failed number=" & CStr(Err.Number)
    Err.Raise Err.Number, "mdlPersonnelHistoryCenter.PreparePersonnelHistoryCorrectionFromCenter", Err.Description
End Sub

Private Function FindEmployeeID(ByVal query As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim matches As Long
    Dim candidate As String

    If Len(query) = 0 Then Err.Raise vbObjectError + 750, "mdlPersonnelHistoryCenter", "Enter EmployeeID, personal number, table number, or exact FIO."
    Set ws = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 1).Value), query, vbTextCompare) = 0 _
            Or StrComp(CellText(ws.Cells(rowNum, 2).Value), query, vbTextCompare) = 0 _
            Or StrComp(CellText(ws.Cells(rowNum, 3).Value), query, vbTextCompare) = 0 _
            Or StrComp(CellText(ws.Cells(rowNum, 4).Value), query, vbTextCompare) = 0 Then
            matches = matches + 1
            candidate = CellText(ws.Cells(rowNum, 1).Value)
        End If
    Next rowNum
    If matches = 0 Then Err.Raise vbObjectError + 751, "mdlPersonnelHistoryCenter", "No employee matches the search value."
    If matches > 1 Then Err.Raise vbObjectError + 752, "mdlPersonnelHistoryCenter", "More than one employee matches the search value. Use EmployeeID."
    FindEmployeeID = candidate
End Function

Private Function FindEmployeeRow(ByVal employeeID As String) As Long
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 1).Value), employeeID, vbTextCompare) = 0 Then FindEmployeeRow = rowNum: Exit Function
    Next rowNum
End Function

Private Function EventBelongsToEmployee(ByVal eventID As String, ByVal employeeID As String) As Boolean
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(EVENTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 1).Value), eventID, vbTextCompare) = 0 Then
            EventBelongsToEmployee = (StrComp(CellText(ws.Cells(rowNum, 2).Value), employeeID, vbTextCompare) = 0)
            Exit Function
        End If
    Next rowNum
End Function

Private Sub AppendEvents(ByRef resultText As String, ByVal employeeID As String, ByRef eventCount As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim count As Long
    Dim ids() As String
    Dim dates() As Double
    Dim lines() As String
    Dim i As Long
    Dim j As Long
    Dim swapText As String
    Dim swapDate As Double

    Set ws = ThisWorkbook.Worksheets(EVENTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 2).Value), employeeID, vbTextCompare) = 0 Then
            count = count + 1
            ReDim Preserve ids(1 To count)
            ReDim Preserve dates(1 To count)
            ReDim Preserve lines(1 To count)
            ids(count) = CellText(ws.Cells(rowNum, 1).Value)
            dates(count) = DateSortValue(ws.Cells(rowNum, 4).Value)
            lines(count) = "EVENT | id=" & ids(count) & " | date=" & DateText(ws.Cells(rowNum, 4).Value) & " | effective=" & DateText(ws.Cells(rowNum, 5).Value) & " | type=" & CellText(ws.Cells(rowNum, 3).Value) & " | status=" & CellText(ws.Cells(rowNum, 6).Value) & " | order=" & CellText(ws.Cells(rowNum, 9).Value)
        End If
    Next rowNum
    For i = 1 To count - 1
        For j = i + 1 To count
            If dates(j) < dates(i) Or (dates(j) = dates(i) And StrComp(ids(j), ids(i), vbTextCompare) < 0) Then
                swapDate = dates(i): dates(i) = dates(j): dates(j) = swapDate
                swapText = ids(i): ids(i) = ids(j): ids(j) = swapText
                swapText = lines(i): lines(i) = lines(j): lines(j) = swapText
            End If
        Next j
    Next i
    For i = 1 To count: AppendLine resultText, lines(i): Next i
    eventCount = count
    If count = 0 Then AppendLine resultText, "No saved personnel events."
End Sub

Private Sub AppendSnapshots(ByRef resultText As String, ByVal employeeID As String, ByRef itemCount As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(SNAPSHOTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 4).Value), employeeID, vbTextCompare) = 0 Then
            AppendLine resultText, "SNAPSHOT | id=" & CellText(ws.Cells(rowNum, 1).Value) & " | event=" & CellText(ws.Cells(rowNum, 2).Value) & " | kind=" & CellText(ws.Cells(rowNum, 3).Value) & " | date=" & DateText(ws.Cells(rowNum, 16).Value) & " | position=" & CellText(ws.Cells(rowNum, 6).Value)
            itemCount = itemCount + 1
        End If
    Next rowNum
    If itemCount = 0 Then AppendLine resultText, "No state snapshots."
End Sub

Private Sub AppendAssignments(ByRef resultText As String, ByVal employeeID As String, ByRef itemCount As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(ASSIGNMENTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 2).Value), employeeID, vbTextCompare) = 0 Then
            AppendLine resultText, "ASSIGNMENT | id=" & CellText(ws.Cells(rowNum, 1).Value) & " | event=" & CellText(ws.Cells(rowNum, 3).Value) & " | type=" & CellText(ws.Cells(rowNum, 4).Value) & " | code=" & CellText(ws.Cells(rowNum, 5).Value) & " | start=" & DateText(ws.Cells(rowNum, 9).Value) & " | end=" & DateText(ws.Cells(rowNum, 10).Value) & " | status=" & CellText(ws.Cells(rowNum, 11).Value)
            itemCount = itemCount + 1
        End If
    Next rowNum
    If itemCount = 0 Then AppendLine resultText, "No payment assignments."
End Sub

Private Sub AppendDocuments(ByRef resultText As String, ByVal employeeID As String, ByVal selectedEventID As String, ByRef itemCount As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim eventID As String
    Set ws = ThisWorkbook.Worksheets(DOCUMENTS_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        eventID = CellText(ws.Cells(rowNum, 2).Value)
        If EventBelongsToEmployee(eventID, employeeID) And (Len(Trim$(selectedEventID)) = 0 Or StrComp(eventID, Trim$(selectedEventID), vbTextCompare) = 0) Then
            AppendLine resultText, "DOCUMENT | id=" & CellText(ws.Cells(rowNum, 1).Value) & " | event=" & eventID & " | type=" & CellText(ws.Cells(rowNum, 3).Value) & " | number=" & CellText(ws.Cells(rowNum, 4).Value) & " | date=" & DateText(ws.Cells(rowNum, 5).Value) & " | status=" & CellText(ws.Cells(rowNum, 10).Value) & " | path=" & CellText(ws.Cells(rowNum, 6).Value)
            itemCount = itemCount + 1
        End If
    Next rowNum
    If itemCount = 0 Then AppendLine resultText, "No registered documents."
End Sub

Private Sub AppendStaffSync(ByRef resultText As String, ByVal employeeID As String, ByRef itemCount As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(STAFF_SYNC_LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then AppendLine resultText, "No staff state synchronizations.": Exit Sub
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 2).Value), employeeID, vbTextCompare) = 0 Then
            AppendLine resultText, "STAFF_SYNC | id=" & CellText(ws.Cells(rowNum, 1).Value) & " | staff=" & CellText(ws.Cells(rowNum, 3).Value) & " | status=" & CellText(ws.Cells(rowNum, 4).Value) & " | fields=" & CellText(ws.Cells(rowNum, 5).Value) & " | date=" & DateText(ws.Cells(rowNum, 7).Value)
            itemCount = itemCount + 1
        End If
    Next rowNum
    If itemCount = 0 Then AppendLine resultText, "No staff state synchronizations."
End Sub

Private Function CurrentStateLine(ByVal employeeID As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(CURRENT_STATE_SHEET)
    lastRow = LastDataRow(ws)
    For rowNum = 2 To lastRow
        If StrComp(CellText(ws.Cells(rowNum, 1).Value), employeeID, vbTextCompare) = 0 Then
            CurrentStateLine = "rank=" & CellText(ws.Cells(rowNum, 2).Value) & " | position=" & CellText(ws.Cells(rowNum, 4).Value) & " | section=" & CellText(ws.Cells(rowNum, 5).Value) & " | unit=" & CellText(ws.Cells(rowNum, 6).Value) & " | state_date=" & DateText(ws.Cells(rowNum, 14).Value) & " | last_event=" & CellText(ws.Cells(rowNum, 16).Value)
            Exit Function
        End If
    Next rowNum
    CurrentStateLine = "not found"
End Function

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    On Error Resume Next
    Set found = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False)
    On Error GoTo 0
    If found Is Nothing Then LastDataRow = 1 Else LastDataRow = found.Row
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    CellText = Trim$(CStr(value))
End Function

Private Function DateText(ByVal value As Variant) As String
    If IsDate(value) Then DateText = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss") Else DateText = CellText(value)
End Function

Private Function DateSortValue(ByVal value As Variant) As Double
    If IsDate(value) Then DateSortValue = CDbl(CDate(value)) Else DateSortValue = 0
End Function

Private Sub AppendLine(ByRef resultText As String, ByVal lineText As String)
    If Len(resultText) > 0 Then resultText = resultText & vbCrLf
    resultText = resultText & lineText
End Sub
