Attribute VB_Name = "mdlStaffLinking"
Option Explicit

' Safe linking of internal employee cards to a Staff export. Candidates are
' based only on exact personal-number matches and require explicit approval.

Private Const REVIEW_SHEET As String = "StaffLinkReview"
Private Const EMPLOYEES_SHEET As String = "Employees"
Private Const CURRENT_STATE_SHEET As String = "EmployeeCurrentState"
Private Const SYNC_LOG_SHEET As String = "StaffStateSyncLog"
Private Const REVIEW_HEADER_ROW As Long = 4
Private Const REVIEW_DATA_ROW As Long = 5
Private Const STAFF_ROW_PREFIX As String = "STAFF_ROW:"

Public Sub OpenStaffLinkReview()
    EnsureStaffLinkReviewSheet
    ThisWorkbook.Worksheets(REVIEW_SHEET).Activate
End Sub

Public Sub BuildStaffLinkCandidates(Optional ByVal showMessage As Boolean = True)
    Dim review As Worksheet
    Dim employees As Worksheet
    Dim staff As Worksheet
    Dim personalColumn As Long
    Dim rankColumn As Long
    Dim fioColumn As Long
    Dim positionColumn As Long
    Dim unitColumn As Long
    Dim employeeRow As Long
    Dim outputRow As Long
    Dim lastEmployeeRow As Long
    Dim personalNumber As String
    Dim matchCount As Long
    Dim matchedStaffRow As Long
    Dim candidateStatus As String
    Dim currentStatus As String

    EnsureStaffLinkReviewSheet
    Set review = ThisWorkbook.Worksheets(REVIEW_SHEET)
    Set employees = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    Set staff = GetStaffWorksheet()
    If staff Is Nothing Then Err.Raise vbObjectError + 760, "mdlStaffLinking", "Staff worksheet was not found."
    If Not FindColumnNumbers(staff, personalColumn, rankColumn, fioColumn, positionColumn, unitColumn) Then Err.Raise vbObjectError + 761, "mdlStaffLinking", "Staff worksheet does not contain required identification columns."

    review.Range("A" & REVIEW_DATA_ROW & ":J2000").ClearContents
    lastEmployeeRow = employees.Cells(employees.Rows.Count, 1).End(xlUp).Row
    outputRow = REVIEW_DATA_ROW
    For employeeRow = 2 To lastEmployeeRow
        personalNumber = LinkText(employees.Cells(employeeRow, 3).Value)
        currentStatus = UCase$(LinkText(employees.Cells(employeeRow, 6).Value))
        matchCount = 0
        matchedStaffRow = 0
        If currentStatus = "LINKED" Then
            candidateStatus = "LINKED"
            matchedStaffRow = ParseStaffRow(LinkText(employees.Cells(employeeRow, 7).Value))
            If personalNumber <> "" Then matchCount = CountStaffMatches(staff, personalColumn, personalNumber, matchedStaffRow)
        ElseIf personalNumber = "" Then
            candidateStatus = "MANUAL_ONLY"
        Else
            matchCount = CountStaffMatches(staff, personalColumn, personalNumber, matchedStaffRow)
            If matchCount = 1 Then
                candidateStatus = "CANDIDATE"
            ElseIf matchCount = 0 Then
                candidateStatus = "MANUAL_ONLY"
            Else
                candidateStatus = "CONFLICT"
            End If
        End If

        If candidateStatus <> "LINKED" Then UpdateEmployeeLinkStatus employees, employeeRow, candidateStatus, IIf(candidateStatus = "CANDIDATE", STAFF_ROW_PREFIX & CStr(matchedStaffRow), "")
        WriteReviewRow review, outputRow, employees, employeeRow, staff, matchedStaffRow, candidateStatus, matchCount
        outputRow = outputRow + 1
    Next employeeRow

    review.Columns("A:J").EntireColumn.AutoFit
    review.Activate
    If showMessage Then MsgBox "Staff link candidates were built. Confirm only rows marked CANDIDATE.", vbInformation
End Sub

Public Sub ConfirmStaffLinkSelections(Optional ByVal showMessage As Boolean = True)
    ApplyStaffLinkSelections "CONFIRM", showMessage
End Sub

Public Sub RejectStaffLinkSelections(Optional ByVal showMessage As Boolean = True)
    ApplyStaffLinkSelections "REJECT", showMessage
End Sub

Public Sub SyncConfirmedStaffState(Optional ByVal showMessage As Boolean = True)
    Dim review As Worksheet
    Dim employees As Worksheet
    Dim staff As Worksheet
    Dim personalColumn As Long
    Dim rankColumn As Long
    Dim fioColumn As Long
    Dim positionColumn As Long
    Dim unitColumn As Long
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeRow As Long
    Dim employeeID As String
    Dim personalNumber As String
    Dim staffRow As Long
    Dim matchedStaffRow As Long
    Dim matchCount As Long
    Dim fieldsChanged As String
    Dim syncStatus As String
    Dim processed As Long

    EnsureStaffLinkReviewSheet
    Set review = ThisWorkbook.Worksheets(REVIEW_SHEET)
    Set employees = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    Set staff = GetStaffWorksheet()
    If staff Is Nothing Then Err.Raise vbObjectError + 764, "mdlStaffLinking", "Staff worksheet was not found."
    If Not FindColumnNumbers(staff, personalColumn, rankColumn, fioColumn, positionColumn, unitColumn) Then Err.Raise vbObjectError + 765, "mdlStaffLinking", "Staff worksheet does not contain required synchronization columns."

    lastRow = review.Cells(review.Rows.Count, 1).End(xlUp).Row
    For rowNum = REVIEW_DATA_ROW To lastRow
        If UCase$(LinkText(review.Cells(rowNum, 9).Value)) = "SYNC" Then
            employeeID = LinkText(review.Cells(rowNum, 1).Value)
            staffRow = ParseStaffRow(LinkText(review.Cells(rowNum, 5).Value))
            employeeRow = FindEmployeeRow(employees, employeeID)
            syncStatus = ""
            fieldsChanged = ""

            If employeeRow = 0 Then
                syncStatus = "FAILED_EMPLOYEE_NOT_FOUND"
            ElseIf UCase$(LinkText(employees.Cells(employeeRow, 6).Value)) <> "LINKED" Then
                syncStatus = "FAILED_NOT_LINKED"
            ElseIf ParseStaffRow(LinkText(employees.Cells(employeeRow, 7).Value)) <> staffRow Or staffRow = 0 Then
                syncStatus = "FAILED_REFERENCE_MISMATCH"
            Else
                personalNumber = LinkText(employees.Cells(employeeRow, 3).Value)
                matchCount = CountStaffMatches(staff, personalColumn, personalNumber, matchedStaffRow)
                If personalNumber = "" Or matchCount <> 1 Or matchedStaffRow <> staffRow Then
                    syncStatus = "FAILED_MATCH_CHANGED"
                Else
                    fieldsChanged = SyncCurrentStateFromStaff(employeeID, staff, staffRow, rankColumn, positionColumn, unitColumn)
                    If Left$(fieldsChanged, 7) = "FAILED_" Then
                        syncStatus = fieldsChanged
                        fieldsChanged = ""
                    ElseIf fieldsChanged = "" Then
                        syncStatus = "NO_CHANGE"
                    Else
                        syncStatus = "SYNCED"
                        processed = processed + 1
                    End If
                End If
            End If

            WriteStateSyncLog employeeID, STAFF_ROW_PREFIX & CStr(staffRow), syncStatus, fieldsChanged
            review.Cells(rowNum, 10).Value = syncStatus & IIf(fieldsChanged <> "", ": " & fieldsChanged, "")
        End If
    Next rowNum

    If showMessage Then MsgBox CStr(processed) & " linked state synchronization(s) completed.", vbInformation
End Sub

Private Sub EnsureStaffLinkReviewSheet()
    Dim ws As Worksheet
    Dim headers As Variant
    Dim index As Long

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REVIEW_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = REVIEW_SHEET
    End If

    If LinkText(ws.Cells(1, 1).Value) = "" Then
        ws.Cells(1, 1).Value = "STAFF LINK REVIEW"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(2, 1).Value = "Build candidates first. Use CONFIRM or REJECT in column I, then run the matching macro. Linking is never automatic."
    End If
    headers = Array("EmployeeID", "FIO", "PersonalNumber", "CurrentStatus", "CandidateReference", "StaffFIO", "StaffPosition", "MatchCount", "Action", "Result")
    For index = LBound(headers) To UBound(headers)
        ws.Cells(REVIEW_HEADER_ROW, index + 1).Value = headers(index)
        ws.Cells(REVIEW_HEADER_ROW, index + 1).Font.Bold = True
        ws.Cells(REVIEW_HEADER_ROW, index + 1).Interior.Color = RGB(217, 225, 242)
    Next index
    EnsureStateSyncLogSheet
End Sub

Private Function SyncCurrentStateFromStaff(ByVal employeeID As String, ByVal staff As Worksheet, ByVal staffRow As Long, ByVal rankColumn As Long, ByVal positionColumn As Long, ByVal unitColumn As Long) As String
    Dim stateSheet As Worksheet
    Dim stateRow As Long
    Dim fieldsChanged As String

    Set stateSheet = ThisWorkbook.Worksheets(CURRENT_STATE_SHEET)
    stateRow = FindCurrentStateRow(stateSheet, employeeID)
    If stateRow = 0 Then
        SyncCurrentStateFromStaff = "FAILED_CURRENT_STATE_NOT_FOUND"
        Exit Function
    End If

    UpdateStateField stateSheet.Cells(stateRow, 2), staff.Cells(staffRow, rankColumn).Value, "Rank", fieldsChanged
    UpdateStateField stateSheet.Cells(stateRow, 4), staff.Cells(staffRow, positionColumn).Value, "Position", fieldsChanged
    UpdateStateField stateSheet.Cells(stateRow, 6), staff.Cells(staffRow, unitColumn).Value, "MilitaryUnit", fieldsChanged
    SyncCurrentStateFromStaff = fieldsChanged
End Function

Private Function FindCurrentStateRow(ByVal stateSheet As Worksheet, ByVal employeeID As String) As Long
    Dim rowNum As Long
    Dim lastRow As Long

    lastRow = stateSheet.Cells(stateSheet.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(LinkText(stateSheet.Cells(rowNum, 1).Value), employeeID, vbTextCompare) = 0 Then
            FindCurrentStateRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Sub UpdateStateField(ByVal targetCell As Range, ByVal sourceValue As Variant, ByVal fieldName As String, ByRef fieldsChanged As String)
    If StrComp(LinkText(targetCell.Value), LinkText(sourceValue), vbTextCompare) = 0 Then Exit Sub
    targetCell.Value = sourceValue
    If fieldsChanged <> "" Then fieldsChanged = fieldsChanged & ", "
    fieldsChanged = fieldsChanged & fieldName
End Sub

Private Sub EnsureStateSyncLogSheet()
    Dim ws As Worksheet
    Dim headers As Variant
    Dim index As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SYNC_LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SYNC_LOG_SHEET
    End If
    headers = Array("SyncID", "EmployeeID", "StaffReference", "Status", "FieldsChanged", "OperatorName", "SyncedAt")
    For index = LBound(headers) To UBound(headers)
        If LinkText(ws.Cells(1, index + 1).Value) = "" Then ws.Cells(1, index + 1).Value = headers(index)
        ws.Cells(1, index + 1).Font.Bold = True
    Next index
    ws.Rows(1).Interior.Color = RGB(217, 225, 242)
End Sub

Private Sub WriteStateSyncLog(ByVal employeeID As String, ByVal staffReference As String, ByVal statusValue As String, ByVal fieldsChanged As String)
    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Worksheets(SYNC_LOG_SHEET)
    rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If rowNum < 2 Then rowNum = 2
    ws.Cells(rowNum, 1).Value = "SSY-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(rowNum)
    ws.Cells(rowNum, 2).Value = employeeID
    ws.Cells(rowNum, 3).Value = staffReference
    ws.Cells(rowNum, 4).Value = statusValue
    ws.Cells(rowNum, 5).Value = fieldsChanged
    ws.Cells(rowNum, 6).Value = Application.UserName
    ws.Cells(rowNum, 7).Value = Now
End Sub

Private Sub WriteReviewRow(ByVal review As Worksheet, ByVal outputRow As Long, ByVal employees As Worksheet, ByVal employeeRow As Long, ByVal staff As Worksheet, ByVal staffRow As Long, ByVal candidateStatus As String, ByVal matchCount As Long)
    Dim staffFio As String
    Dim staffPosition As String
    Dim personalColumn As Long
    Dim rankColumn As Long
    Dim fioColumn As Long
    Dim positionColumn As Long
    Dim unitColumn As Long

    If staffRow > 0 Then
        If Not FindColumnNumbers(staff, personalColumn, rankColumn, fioColumn, positionColumn, unitColumn) Then Exit Sub
        staffFio = LinkText(staff.Cells(staffRow, fioColumn).Value)
        staffPosition = LinkText(staff.Cells(staffRow, positionColumn).Value)
    End If
    review.Cells(outputRow, 1).Value = employees.Cells(employeeRow, 1).Value
    review.Cells(outputRow, 2).Value = employees.Cells(employeeRow, 2).Value
    review.Cells(outputRow, 3).Value = employees.Cells(employeeRow, 3).Value
    review.Cells(outputRow, 4).Value = candidateStatus
    If staffRow > 0 Then review.Cells(outputRow, 5).Value = STAFF_ROW_PREFIX & CStr(staffRow)
    review.Cells(outputRow, 6).Value = staffFio
    review.Cells(outputRow, 7).Value = staffPosition
    review.Cells(outputRow, 8).Value = matchCount
End Sub

Private Sub ApplyStaffLinkSelections(ByVal requiredAction As String, ByVal showMessage As Boolean)
    Dim review As Worksheet
    Dim employees As Worksheet
    Dim staff As Worksheet
    Dim personalColumn As Long
    Dim rankColumn As Long
    Dim fioColumn As Long
    Dim positionColumn As Long
    Dim unitColumn As Long
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeRow As Long
    Dim staffRow As Long
    Dim matchCount As Long
    Dim matchedStaffRow As Long
    Dim employeeID As String
    Dim personalNumber As String
    Dim actionValue As String
    Dim processed As Long

    EnsureStaffLinkReviewSheet
    Set review = ThisWorkbook.Worksheets(REVIEW_SHEET)
    Set employees = ThisWorkbook.Worksheets(EMPLOYEES_SHEET)
    Set staff = GetStaffWorksheet()
    If staff Is Nothing Then Err.Raise vbObjectError + 762, "mdlStaffLinking", "Staff worksheet was not found."
    If Not FindColumnNumbers(staff, personalColumn, rankColumn, fioColumn, positionColumn, unitColumn) Then Err.Raise vbObjectError + 763, "mdlStaffLinking", "Staff worksheet does not contain required identification columns."

    lastRow = review.Cells(review.Rows.Count, 1).End(xlUp).Row
    For rowNum = REVIEW_DATA_ROW To lastRow
        actionValue = UCase$(LinkText(review.Cells(rowNum, 9).Value))
        If actionValue = requiredAction Then
            employeeID = LinkText(review.Cells(rowNum, 1).Value)
            employeeRow = FindEmployeeRow(employees, employeeID)
            If employeeRow = 0 Then
                review.Cells(rowNum, 10).Value = "Employee not found"
            ElseIf requiredAction = "REJECT" Then
                UpdateEmployeeLinkStatus employees, employeeRow, "MANUAL_ONLY", ""
                review.Cells(rowNum, 4).Value = "MANUAL_ONLY"
                review.Cells(rowNum, 5).ClearContents
                review.Cells(rowNum, 10).Value = "Rejected"
                processed = processed + 1
            ElseIf UCase$(LinkText(review.Cells(rowNum, 4).Value)) <> "CANDIDATE" Then
                review.Cells(rowNum, 10).Value = "Only CANDIDATE can be confirmed"
            Else
                personalNumber = LinkText(employees.Cells(employeeRow, 3).Value)
                staffRow = ParseStaffRow(LinkText(review.Cells(rowNum, 5).Value))
                matchCount = CountStaffMatches(staff, personalColumn, personalNumber, matchedStaffRow)
                If personalNumber = "" Or matchCount <> 1 Or matchedStaffRow <> staffRow Then
                    review.Cells(rowNum, 10).Value = "Candidate no longer valid; rebuild review"
                Else
                    UpdateEmployeeLinkStatus employees, employeeRow, "LINKED", STAFF_ROW_PREFIX & CStr(staffRow)
                    review.Cells(rowNum, 4).Value = "LINKED"
                    review.Cells(rowNum, 10).Value = "Linked"
                    processed = processed + 1
                End If
            End If
        End If
    Next rowNum
    If Not showMessage Then Exit Sub
    If processed = 0 Then
        MsgBox "No valid " & requiredAction & " actions were processed.", vbExclamation
    Else
        MsgBox CStr(processed) & " staff-link action(s) processed.", vbInformation
    End If
End Sub

Private Function CountStaffMatches(ByVal staff As Worksheet, ByVal personalColumn As Long, ByVal personalNumber As String, ByRef matchedRow As Long) As Long
    Dim rowNum As Long
    Dim lastRow As Long

    lastRow = staff.Cells(staff.Rows.Count, personalColumn).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(LinkText(staff.Cells(rowNum, personalColumn).Value), personalNumber, vbTextCompare) = 0 Then
            CountStaffMatches = CountStaffMatches + 1
            matchedRow = rowNum
        End If
    Next rowNum
End Function

Private Function FindEmployeeRow(ByVal employees As Worksheet, ByVal employeeID As String) As Long
    Dim rowNum As Long
    Dim lastRow As Long

    lastRow = employees.Cells(employees.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(LinkText(employees.Cells(rowNum, 1).Value), employeeID, vbTextCompare) = 0 Then
            FindEmployeeRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Sub UpdateEmployeeLinkStatus(ByVal employees As Worksheet, ByVal employeeRow As Long, ByVal linkStatus As String, ByVal staffReference As String)
    employees.Cells(employeeRow, 6).Value = linkStatus
    employees.Cells(employeeRow, 7).Value = staffReference
    employees.Cells(employeeRow, 9).Value = Now
End Sub

Private Function ParseStaffRow(ByVal referenceValue As String) As Long
    If Left$(referenceValue, Len(STAFF_ROW_PREFIX)) <> STAFF_ROW_PREFIX Then Exit Function
    If IsNumeric(Mid$(referenceValue, Len(STAFF_ROW_PREFIX) + 1)) Then ParseStaffRow = CLng(Mid$(referenceValue, Len(STAFF_ROW_PREFIX) + 1))
End Function

Private Function LinkText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    LinkText = Trim$(CStr(value))
End Function
