Attribute VB_Name = "mdlPersonnelHistory"
Option Explicit

' Read-only employee history view. It is deliberately derived from the
' append-only personnel-event ledger and never changes saved snapshots.

Private Const HISTORY_SHEET As String = "PersonnelHistory"
Private Const EMPLOYEES_SHEET As String = "Employees"
Private Const CURRENT_STATE_SHEET As String = "EmployeeCurrentState"
Private Const EVENTS_SHEET As String = "PersonnelEvents"
Private Const SNAPSHOTS_SHEET As String = "PersonnelStateSnapshots"
Private Const ASSIGNMENTS_SHEET As String = "PaymentAssignments"
Private Const DOCUMENTS_SHEET As String = "DocumentRegistry"
Private Const STAFF_SYNC_LOG_SHEET As String = "StaffStateSyncLog"

Public Sub OpenPersonnelHistory()
    EnsurePersonnelHistorySheet
    ThisWorkbook.Worksheets(HISTORY_SHEET).Activate
End Sub

Public Sub SearchPersonnelHistory()
    Dim ws As Worksheet
    Dim query As String
    Dim employeeID As String

    EnsurePersonnelHistorySheet
    Set ws = ThisWorkbook.Worksheets(HISTORY_SHEET)
    query = HistoryText(ws.Cells(3, 2).Value)
    If query = "" Then
        MsgBox "Enter EmployeeID, personal number, or exact FIO in B3.", vbExclamation
        Exit Sub
    End If

    employeeID = FindEmployeeID(query)
    ws.Cells(4, 2).Value = employeeID
    RefreshPersonnelHistory
End Sub

Public Sub RefreshPersonnelHistory()
    Dim ws As Worksheet
    Dim employeeID As String

    EnsurePersonnelHistorySheet
    Set ws = ThisWorkbook.Worksheets(HISTORY_SHEET)
    employeeID = HistoryText(ws.Cells(4, 2).Value)
    If employeeID = "" Then
        MsgBox "Select an employee first or search in B3.", vbExclamation
        Exit Sub
    End If
    If FindEmployeeRow(employeeID) = 0 Then
        MsgBox "EmployeeID was not found: " & employeeID, vbExclamation
        Exit Sub
    End If

    RenderPersonnelHistory ws, employeeID
End Sub

Public Sub ExportPersonnelHistoryEvent()
    Dim eventID As String
    Dim outputPath As String

    eventID = ""
    On Error Resume Next
    eventID = HistoryText(ActiveCell.Value)
    On Error GoTo 0
    If Left$(eventID, 4) <> "EVT-" Then eventID = Trim$(InputBox("Enter EventID to export from saved snapshots:", "Personnel order export"))
    If eventID = "" Then Exit Sub

    On Error GoTo ExportError
    outputPath = mdlPersonnelEventOrderExport.ExportPersonnelEventOrder(eventID)
    MsgBox "Personnel order was exported to:" & vbCrLf & outputPath, vbInformation
    Exit Sub
ExportError:
    MsgBox "Personnel order export failed:" & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub PreparePersonnelHistoryCorrection(Optional ByVal showMessage As Boolean = True)
    Dim historySheet As Worksheet
    Dim employeeID As String
    Dim eventID As String

    EnsurePersonnelHistorySheet
    Set historySheet = ThisWorkbook.Worksheets(HISTORY_SHEET)
    employeeID = HistoryText(historySheet.Cells(4, 2).Value)
    If employeeID = "" Then
        MsgBox "Search for an employee before preparing a correction.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    eventID = HistoryText(ActiveCell.Value)
    On Error GoTo 0
    If Left$(eventID, 4) <> "EVT-" Then eventID = Trim$(InputBox("Select an EventID in the history or enter it explicitly:", "Prepare correction"))
    If eventID = "" Then Exit Sub

    On Error GoTo PrepareError
    mdlPersonnelEvents.PreparePersonnelEventCorrection employeeID, eventID
    If showMessage Then MsgBox "Correction form is prepared. Review event type, dates, grounds and payment conditions before saving.", vbInformation
    Exit Sub
PrepareError:
    If showMessage Then MsgBox "Correction preparation failed:" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub EnsurePersonnelHistorySheet()
    Dim ws As Worksheet

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HISTORY_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = HISTORY_SHEET
    End If

    If HistoryText(ws.Cells(1, 1).Value) = "" Then
        ws.Cells(1, 1).Value = "PERSONNEL HISTORY"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(2, 1).Value = "1. Enter EmployeeID, personal number, or exact FIO in B3. 2. Run SearchPersonnelHistory. 3. Select an EventID to export or prepare a correction."
        ws.Cells(3, 1).Value = "Search"
        ws.Cells(4, 1).Value = "EmployeeID"
        ws.Columns("A:M").ColumnWidth = 18
        ws.Columns("B").ColumnWidth = 32
    End If
End Sub

Private Function FindEmployeeID(ByVal query As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim matches As Long
    Dim matchedID As String

    Set ws = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(HistoryText(ws.Cells(rowNum, 1).Value), query, vbTextCompare) = 0 _
            Or StrComp(HistoryText(ws.Cells(rowNum, 2).Value), query, vbTextCompare) = 0 _
            Or StrComp(HistoryText(ws.Cells(rowNum, 3).Value), query, vbTextCompare) = 0 Then
            matches = matches + 1
            matchedID = HistoryText(ws.Cells(rowNum, 1).Value)
        End If
    Next rowNum

    If matches = 0 Then Err.Raise vbObjectError + 740, "mdlPersonnelHistory", "No employee matches the search value."
    If matches > 1 Then Err.Raise vbObjectError + 741, "mdlPersonnelHistory", "More than one employee matches the search value. Use EmployeeID."
    FindEmployeeID = matchedID
End Function

Private Function FindEmployeeRow(ByVal employeeID As String) As Long
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(HistoryText(ws.Cells(rowNum, 1).Value), employeeID, vbTextCompare) = 0 Then
            FindEmployeeRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Sub RenderPersonnelHistory(ByVal historySheet As Worksheet, ByVal employeeID As String)
    Dim employeeSheet As Worksheet
    Dim employeeRow As Long
    Dim outputRow As Long
    Dim eventIDs As Object

    Set employeeSheet = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    employeeRow = FindEmployeeRow(employeeID)
    historySheet.Range("A6:O2000").ClearContents

    historySheet.Cells(6, 1).Value = "Employee card"
    historySheet.Cells(6, 1).Font.Bold = True
    WriteRow historySheet, 7, Array("EmployeeID", "FIO", "PersonalNumber", "TableNumber", "SourceMode", "StaffLinkStatus", "StaffReference", "IsActive")
    WriteRow historySheet, 8, Array( _
        employeeSheet.Cells(employeeRow, 1).Value, employeeSheet.Cells(employeeRow, 2).Value, employeeSheet.Cells(employeeRow, 3).Value, employeeSheet.Cells(employeeRow, 4).Value, _
        employeeSheet.Cells(employeeRow, 5).Value, employeeSheet.Cells(employeeRow, 6).Value, employeeSheet.Cells(employeeRow, 7).Value, employeeSheet.Cells(employeeRow, 10).Value)
    FormatHeader historySheet.Range("A7:H7")

    outputRow = 10
    historySheet.Cells(outputRow, 1).Value = "Current state"
    historySheet.Cells(outputRow, 1).Font.Bold = True
    WriteCurrentState historySheet, outputRow + 1, employeeID

    outputRow = outputRow + 4
    historySheet.Cells(outputRow, 1).Value = "Personnel events"
    historySheet.Cells(outputRow, 1).Font.Bold = True
    WriteRow historySheet, outputRow + 1, Array("EventID", "EventDate", "EffectiveDate", "EventType", "Status", "OrderReference", "BasisText", "BeforePosition", "AfterPosition")
    FormatHeader historySheet.Range(historySheet.Cells(outputRow + 1, 1), historySheet.Cells(outputRow + 1, 9))
    Set eventIDs = WriteEvents(historySheet, outputRow + 2, employeeID)

    outputRow = historySheet.Cells(historySheet.Rows.Count, 1).End(xlUp).Row + 2
    historySheet.Cells(outputRow, 1).Value = "Payment assignments"
    historySheet.Cells(outputRow, 1).Font.Bold = True
    WriteRow historySheet, outputRow + 1, Array("AssignmentID", "EventID", "PaymentType", "PaymentCode", "AmountKind", "AmountValue", "StartDate", "EndDate", "Status", "TerminationEventID", "ActPoint", "OriginalAmount", "AppliedAmount", "CapGroup", "Explanation")
    FormatHeader historySheet.Range(historySheet.Cells(outputRow + 1, 1), historySheet.Cells(outputRow + 1, 15))
    WriteAssignments historySheet, outputRow + 2, employeeID

    outputRow = historySheet.Cells(historySheet.Rows.Count, 1).End(xlUp).Row + 2
    historySheet.Cells(outputRow, 1).Value = "Generated documents"
    historySheet.Cells(outputRow, 1).Font.Bold = True
    WriteRow historySheet, outputRow + 1, Array("DocumentID", "EventID", "DocumentType", "DocumentNumber", "DocumentDate", "FilePath", "TemplateName", "TemplateVersion", "Status", "LastError")
    FormatHeader historySheet.Range(historySheet.Cells(outputRow + 1, 1), historySheet.Cells(outputRow + 1, 10))
    WriteDocuments historySheet, outputRow + 2, eventIDs

    outputRow = historySheet.Cells(historySheet.Rows.Count, 1).End(xlUp).Row + 2
    historySheet.Cells(outputRow, 1).Value = "Staff state synchronization"
    historySheet.Cells(outputRow, 1).Font.Bold = True
    WriteRow historySheet, outputRow + 1, Array("SyncID", "StaffReference", "Status", "FieldsChanged", "OperatorName", "SyncedAt")
    FormatHeader historySheet.Range(historySheet.Cells(outputRow + 1, 1), historySheet.Cells(outputRow + 1, 6))
    WriteStaffStateSynchronizations historySheet, outputRow + 2, employeeID

    historySheet.Columns("A:O").EntireColumn.AutoFit
    historySheet.Activate
End Sub

Private Sub WriteCurrentState(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal employeeID As String)
    Dim stateSheet As Worksheet
    Dim sourceRow As Long
    Dim lastRow As Long

    Set stateSheet = ThisWorkbook.Worksheets(CURRENT_STATE_SHEET)
    lastRow = stateSheet.Cells(stateSheet.Rows.Count, 1).End(xlUp).Row
    For sourceRow = 2 To lastRow
        If HistoryText(stateSheet.Cells(sourceRow, 1).Value) = employeeID Then Exit For
    Next sourceRow
    WriteRow ws, rowNum, Array("Rank", "Position", "Section", "MilitaryUnit", "VUS", "TariffRank", "PositionSalary", "RankSalary", "ServiceCategory", "ContractKind", "ContractBasis", "StateDate", "LastEventID")
    FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 13))
    If sourceRow <= lastRow Then
        WriteRow ws, rowNum + 1, Array(stateSheet.Cells(sourceRow, 2).Value, stateSheet.Cells(sourceRow, 4).Value, stateSheet.Cells(sourceRow, 5).Value, stateSheet.Cells(sourceRow, 6).Value, stateSheet.Cells(sourceRow, 7).Value, stateSheet.Cells(sourceRow, 8).Value, stateSheet.Cells(sourceRow, 9).Value, stateSheet.Cells(sourceRow, 10).Value, stateSheet.Cells(sourceRow, 11).Value, stateSheet.Cells(sourceRow, 12).Value, stateSheet.Cells(sourceRow, 13).Value, stateSheet.Cells(sourceRow, 14).Value, stateSheet.Cells(sourceRow, 16).Value)
    End If
End Sub

Private Function WriteEvents(ByVal ws As Worksheet, ByVal outputRow As Long, ByVal employeeID As String) As Object
    Dim eventSheet As Worksheet
    Dim rowNum As Long
    Dim targetRow As Long
    Dim lastRow As Long
    Dim eventIDs As Object
    Dim beforeSnapshotID As String
    Dim afterSnapshotID As String

    Set eventIDs = CreateObject("Scripting.Dictionary")
    Set eventSheet = ThisWorkbook.Worksheets(EVENTS_SHEET)
    lastRow = eventSheet.Cells(eventSheet.Rows.Count, 1).End(xlUp).Row
    targetRow = outputRow
    For rowNum = 2 To lastRow
        If HistoryText(eventSheet.Cells(rowNum, 2).Value) = employeeID Then
            beforeSnapshotID = HistoryText(eventSheet.Cells(rowNum, 7).Value)
            afterSnapshotID = HistoryText(eventSheet.Cells(rowNum, 8).Value)
            WriteRow ws, targetRow, Array(eventSheet.Cells(rowNum, 1).Value, eventSheet.Cells(rowNum, 4).Value, eventSheet.Cells(rowNum, 5).Value, eventSheet.Cells(rowNum, 3).Value, eventSheet.Cells(rowNum, 6).Value, eventSheet.Cells(rowNum, 9).Value, eventSheet.Cells(rowNum, 10).Value, SnapshotValue(beforeSnapshotID, "Position"), SnapshotValue(afterSnapshotID, "Position"))
            eventIDs.Add HistoryText(eventSheet.Cells(rowNum, 1).Value), True
            targetRow = targetRow + 1
        End If
    Next rowNum
    If targetRow = outputRow Then ws.Cells(targetRow, 1).Value = "No saved personnel events."
    Set WriteEvents = eventIDs
End Function

Private Sub WriteAssignments(ByVal ws As Worksheet, ByVal outputRow As Long, ByVal employeeID As String)
    Dim sourceSheet As Worksheet
    Dim rowNum As Long
    Dim targetRow As Long
    Dim lastRow As Long

    Set sourceSheet = ThisWorkbook.Worksheets(ASSIGNMENTS_SHEET)
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
    targetRow = outputRow
    For rowNum = 2 To lastRow
        If HistoryText(sourceSheet.Cells(rowNum, 2).Value) = employeeID Then
            WriteRow ws, targetRow, Array(sourceSheet.Cells(rowNum, 1).Value, sourceSheet.Cells(rowNum, 3).Value, sourceSheet.Cells(rowNum, 4).Value, sourceSheet.Cells(rowNum, 5).Value, sourceSheet.Cells(rowNum, 6).Value, sourceSheet.Cells(rowNum, 7).Value, sourceSheet.Cells(rowNum, 9).Value, sourceSheet.Cells(rowNum, 10).Value, sourceSheet.Cells(rowNum, 11).Value, sourceSheet.Cells(rowNum, 12).Value, sourceSheet.Cells(rowNum, 14).Value, sourceSheet.Cells(rowNum, 17).Value, sourceSheet.Cells(rowNum, 18).Value, sourceSheet.Cells(rowNum, 19).Value, sourceSheet.Cells(rowNum, 20).Value)
            targetRow = targetRow + 1
        End If
    Next rowNum
    If targetRow = outputRow Then ws.Cells(targetRow, 1).Value = "No payment assignments."
End Sub

Private Sub WriteDocuments(ByVal ws As Worksheet, ByVal outputRow As Long, ByVal eventIDs As Object)
    Dim sourceSheet As Worksheet
    Dim rowNum As Long
    Dim targetRow As Long
    Dim lastRow As Long
    Dim eventID As String

    Set sourceSheet = ThisWorkbook.Worksheets(DOCUMENTS_SHEET)
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
    targetRow = outputRow
    For rowNum = 2 To lastRow
        eventID = HistoryText(sourceSheet.Cells(rowNum, 2).Value)
        If eventIDs.Exists(eventID) Then
            WriteRow ws, targetRow, Array(sourceSheet.Cells(rowNum, 1).Value, eventID, sourceSheet.Cells(rowNum, 3).Value, sourceSheet.Cells(rowNum, 4).Value, sourceSheet.Cells(rowNum, 5).Value, sourceSheet.Cells(rowNum, 6).Value, sourceSheet.Cells(rowNum, 7).Value, sourceSheet.Cells(rowNum, 8).Value, sourceSheet.Cells(rowNum, 10).Value, sourceSheet.Cells(rowNum, 11).Value)
            targetRow = targetRow + 1
        End If
    Next rowNum
    If targetRow = outputRow Then ws.Cells(targetRow, 1).Value = "No generated documents."
End Sub

Private Sub WriteStaffStateSynchronizations(ByVal ws As Worksheet, ByVal outputRow As Long, ByVal employeeID As String)
    Dim sourceSheet As Worksheet
    Dim rowNum As Long
    Dim targetRow As Long
    Dim lastRow As Long

    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Worksheets(STAFF_SYNC_LOG_SHEET)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        ws.Cells(outputRow, 1).Value = "No staff state synchronizations."
        Exit Sub
    End If

    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
    targetRow = outputRow
    For rowNum = 2 To lastRow
        If StrComp(HistoryText(sourceSheet.Cells(rowNum, 2).Value), employeeID, vbTextCompare) = 0 Then
            WriteRow ws, targetRow, Array(sourceSheet.Cells(rowNum, 1).Value, sourceSheet.Cells(rowNum, 3).Value, sourceSheet.Cells(rowNum, 4).Value, sourceSheet.Cells(rowNum, 5).Value, sourceSheet.Cells(rowNum, 6).Value, sourceSheet.Cells(rowNum, 7).Value)
            targetRow = targetRow + 1
        End If
    Next rowNum
    If targetRow = outputRow Then ws.Cells(targetRow, 1).Value = "No staff state synchronizations."
End Sub

Private Function SnapshotValue(ByVal snapshotID As String, ByVal fieldName As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim fieldColumn As Long

    If snapshotID = "" Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SNAPSHOTS_SHEET)
    fieldColumn = FindHeaderColumn(ws, fieldName)
    If fieldColumn = 0 Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If HistoryText(ws.Cells(rowNum, 1).Value) = snapshotID Then
            SnapshotValue = HistoryText(ws.Cells(rowNum, fieldColumn).Value)
            Exit Function
        End If
    Next rowNum
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim colNum As Long
    For colNum = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If StrComp(HistoryText(ws.Cells(1, colNum).Value), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = colNum
            Exit Function
        End If
    Next colNum
End Function

Private Sub WriteRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal values As Variant)
    Dim index As Long
    For index = LBound(values) To UBound(values)
        ws.Cells(rowNum, index + 1).Value = values(index)
    Next index
End Sub

Private Sub FormatHeader(ByVal targetRange As Range)
    targetRange.Font.Bold = True
    targetRange.Interior.Color = RGB(217, 225, 242)
End Sub

Private Function HistoryText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    HistoryText = Trim$(CStr(value))
End Function
