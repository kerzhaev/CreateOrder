Attribute VB_Name = "mdlPersonnelDataIntegrity"
Option Explicit

' P2 read-only integrity scanner. This module must never mutate the workbook.

Private Const SHEET_EMPLOYEES As String = "Employees"
Private Const SHEET_CURRENT_STATE As String = "EmployeeCurrentState"
Private Const SHEET_EVENTS As String = "PersonnelEvents"
Private Const SHEET_SNAPSHOTS As String = "PersonnelStateSnapshots"
Private Const SHEET_ASSIGNMENTS As String = "PaymentAssignments"
Private Const SHEET_DOCUMENTS As String = "DocumentRegistry"
Private Const SHEET_LEGAL_ACTS As String = "LegalActs"
Private Const SHEET_PAYMENT_RULES As String = "PaymentRules"
Private Const SHEET_POSITION_CLASSIFICATION As String = "PositionClassification"
Private Const STAFF_ROW_PREFIX As String = "STAFF_ROW:"

Public Function ScanPersonnelDataIntegrity() As Collection
    Dim findings As Collection
    Dim employees As Object
    Dim states As Object
    Dim events As Object
    Dim snapshots As Object
    Dim assignments As Object
    Dim documents As Object
    Dim legalActs As Object
    Dim employeeRows As Object
    Dim eventRows As Object
    Dim snapshotRows As Object
    Dim assignmentRows As Object
    Dim documentRows As Object
    Dim legalActRows As Object
    Dim excludedEmployees As Object

    On Error GoTo Failed
    Set findings = New Collection
    CheckRequiredSchemas findings

    Set employees = GetIntegritySheet(SHEET_EMPLOYEES)
    Set states = GetIntegritySheet(SHEET_CURRENT_STATE)
    Set events = GetIntegritySheet(SHEET_EVENTS)
    Set snapshots = GetIntegritySheet(SHEET_SNAPSHOTS)
    Set assignments = GetIntegritySheet(SHEET_ASSIGNMENTS)
    Set documents = GetIntegritySheet(SHEET_DOCUMENTS)
    Set legalActs = GetIntegritySheet(SHEET_LEGAL_ACTS)

    Set employeeRows = BuildIdentifierIndex(employees, "EmployeeID", "Employees", "Employee", findings)
    Set eventRows = BuildIdentifierIndex(events, "EventID", "PersonnelEvents", "Event", findings)
    Set snapshotRows = BuildIdentifierIndex(snapshots, "SnapshotID", "PersonnelStateSnapshots", "Snapshot", findings)
    Set assignmentRows = BuildIdentifierIndex(assignments, "AssignmentID", "PaymentAssignments", "Assignment", findings)
    Set documentRows = BuildIdentifierIndex(documents, "DocumentID", "DocumentRegistry", "Document", findings)
    Set legalActRows = BuildIdentifierIndex(legalActs, "ActID", "LegalActs", "LegalAct", findings)

    CheckEventEmployeeLinks events, employeeRows, findings
    CheckEventSnapshotLinks events, eventRows, employeeRows, snapshotRows, snapshots, findings
    CheckSnapshotLinks snapshots, eventRows, employeeRows, findings
    CheckAssignmentLinks assignments, employeeRows, eventRows, legalActRows, findings
    CheckDocumentLinks documents, eventRows, findings
    CheckCurrentStateLinks states, employeeRows, eventRows, findings
    CheckChronology events, eventRows, snapshots, findings
    Set excludedEmployees = BuildExcludedEmployeeIndex(events)
    CheckExclusionAndActivePayments employees, assignments, excludedEmployees, findings
    CheckStaffLinks employees, findings
    CheckLegalAndClassificationReferences assignments, legalActs, legalActRows, findings

    SortFindings findings
    LogIntegrityComplete findings
    Set ScanPersonnelDataIntegrity = findings
    Exit Function

Failed:
    If findings Is Nothing Then Set findings = New Collection
    AddFinding findings, "ERROR", "SCAN", "Workbook", "workbook", "Integrity scan could not be completed.", "Review the workbook structure and rerun the read-only scan."
    Debug.Print "ERROR integrity-scan-failed"
    Set ScanPersonnelDataIntegrity = findings
End Function

Public Function BuildPersonnelDataIntegrityReport() As String
    Dim findings As Collection
    Dim item As Object
    Dim result As String
    Dim errors As Long
    Dim warnings As Long

    Set findings = ScanPersonnelDataIntegrity()
    For Each item In findings
        If UCase$(SafeIntegrityText(item("Severity"))) = "ERROR" Then
            errors = errors + 1
        ElseIf UCase$(SafeIntegrityText(item("Severity"))) = "WARNING" Then
            warnings = warnings + 1
        End If
    Next item

    result = "Integrity scan | findings=" & CStr(findings.Count) & "; errors=" & CStr(errors) & "; warnings=" & CStr(warnings)
    For Each item In findings
        result = result & vbCrLf & SafeIntegrityText(item("Severity")) & " | " & SafeIntegrityText(item("Category")) & " | " & SafeIntegrityText(item("EntityType")) & " | " & SafeIntegrityText(item("EntityID")) & " | " & SafeIntegrityText(item("Message")) & " | " & SafeIntegrityText(item("SuggestedAction"))
    Next item
    BuildPersonnelDataIntegrityReport = result
End Function

Public Function PersonnelDataIntegritySummary() As Object
    Dim summary As Object
    Dim findings As Collection
    Dim item As Object
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim totalRows As Long
    Dim errorCount As Long
    Dim warningCount As Long

    Set summary = CreateObject("Scripting.Dictionary")
    Set findings = ScanPersonnelDataIntegrity()
    sheetNames = Array(SHEET_EMPLOYEES, SHEET_CURRENT_STATE, SHEET_EVENTS, SHEET_SNAPSHOTS, SHEET_ASSIGNMENTS, SHEET_DOCUMENTS, SHEET_LEGAL_ACTS, SHEET_PAYMENT_RULES, SHEET_POSITION_CLASSIFICATION)
    For Each sheetName In sheetNames
        Set ws = GetIntegritySheet(CStr(sheetName))
        If Not ws Is Nothing Then
            summary("Sheets") = CLng(ValueOrZero(summary, "Sheets")) + 1
            totalRows = totalRows + IntegrityDataRowCount(ws)
        End If
    Next sheetName
    For Each item In findings
        If UCase$(SafeIntegrityText(item("Severity"))) = "ERROR" Then errorCount = errorCount + 1
        If UCase$(SafeIntegrityText(item("Severity"))) = "WARNING" Then warningCount = warningCount + 1
    Next item
    summary("Rows") = totalRows
    summary("Findings") = findings.Count
    summary("Errors") = errorCount
    summary("Warnings") = warningCount
    Set PersonnelDataIntegritySummary = summary
End Function

Private Sub CheckRequiredSchemas(ByVal findings As Collection)
    Dim sheetNames As Variant
    Dim headerSets As Variant
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim headers As Variant

    sheetNames = Array(SHEET_EMPLOYEES, SHEET_CURRENT_STATE, SHEET_EVENTS, SHEET_SNAPSHOTS, SHEET_ASSIGNMENTS, SHEET_DOCUMENTS, SHEET_LEGAL_ACTS, SHEET_PAYMENT_RULES, SHEET_POSITION_CLASSIFICATION)
    headerSets = Array( _
        Array("EmployeeID", "FIO", "PersonalNumber", "StaffLinkStatus", "StaffReference", "IsActive"), _
        Array("EmployeeID", "StateDate", "SourceEventID", "LastEventID"), _
        Array("EventID", "EmployeeID", "EventType", "EventDate", "EffectiveDate", "Status", "BeforeSnapshotID", "AfterSnapshotID", "CreatedAt"), _
        Array("SnapshotID", "EventID", "SnapshotKind", "EmployeeID", "StateDate", "CreatedAt"), _
        Array("AssignmentID", "EmployeeID", "EventID", "Status", "StartDate", "EndDate", "ActID"), _
        Array("DocumentID", "EventID", "FilePath", "Status"), _
        Array("ActID", "ActType", "ActNumber", "ActDate", "EffectiveFrom"), _
        Array("RuleID", "PaymentCode", "ActID", "RuleStatus"), _
        Array("ClassificationID", "PositionKey", "ReviewStatus"))

    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = GetIntegritySheet(CStr(sheetNames(i)))
        If ws Is Nothing Then
            AddFinding findings, "ERROR", "SCHEMA", "Worksheet", CStr(sheetNames(i)), "Required worksheet is missing.", "Restore the worksheet from a trusted workbook copy."
        Else
            headers = headerSets(i)
            For j = LBound(headers) To UBound(headers)
                If HeaderColumn(ws, CStr(headers(j))) = 0 Then
                    AddFinding findings, "ERROR", "SCHEMA", "Worksheet", CStr(sheetNames(i)), "Required header is missing.", "Restore the required header without changing existing data."
                End If
            Next j
        End If
    Next i
End Sub

Private Function BuildIdentifierIndex(ByVal ws As Worksheet, ByVal idHeader As String, ByVal sheetName As String, ByVal entityType As String, ByVal findings As Collection) As Object
    Dim index As Object
    Dim idColumn As Long
    Dim rowNum As Long
    Dim lastRow As Long
    Dim valueText As String

    Set index = CreateObject("Scripting.Dictionary")
    index.CompareMode = vbTextCompare
    If ws Is Nothing Then
        Set BuildIdentifierIndex = index
        Exit Function
    End If
    idColumn = HeaderColumn(ws, idHeader)
    If idColumn = 0 Then
        Set BuildIdentifierIndex = index
        Exit Function
    End If

    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        valueText = CellText(ws.Cells(rowNum, idColumn).Value)
        If valueText = "" Then
            If RowHasAnyValue(ws, rowNum) Then AddFinding findings, "ERROR", "IDENTIFIER", entityType, RowToken(sheetName, rowNum), "Identifier is blank.", "Provide a unique identifier or remove the incomplete row."
        ElseIf index.Exists(valueText) Then
            AddFinding findings, "ERROR", "IDENTIFIER", entityType, RowToken(sheetName, rowNum), "Identifier is duplicated.", "Keep one authoritative row and correct the duplicate manually."
        Else
            index.Add valueText, rowNum
        End If
    Next rowNum
    Set BuildIdentifierIndex = index
End Function

Private Sub CheckEventEmployeeLinks(ByVal ws As Worksheet, ByVal employeeRows As Object, ByVal findings As Collection)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeID As String
    If ws Is Nothing Or employeeRows Is Nothing Then Exit Sub
    If HeaderColumn(ws, "EmployeeID") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        If RowHasAnyValue(ws, rowNum) Then
            employeeID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EmployeeID")).Value)
            If employeeID = "" Or Not employeeRows.Exists(employeeID) Then AddFinding findings, "ERROR", "EVENT_EMPLOYEE", "Event", RowToken(SHEET_EVENTS, rowNum), "Event references a missing employee.", "Link the event to an existing employee record."
        End If
    Next rowNum
End Sub

Private Sub CheckEventSnapshotLinks(ByVal ws As Worksheet, ByVal eventRows As Object, ByVal employeeRows As Object, ByVal snapshotRows As Object, ByVal snapshots As Worksheet, ByVal findings As Collection)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim eventID As String
    Dim employeeID As String
    Dim beforeID As String
    Dim afterID As String
    If ws Is Nothing Or eventRows Is Nothing Or employeeRows Is Nothing Or snapshotRows Is Nothing Or snapshots Is Nothing Then Exit Sub
    If HeaderColumn(ws, "EventID") = 0 Or HeaderColumn(ws, "EmployeeID") = 0 Or HeaderColumn(ws, "BeforeSnapshotID") = 0 Or HeaderColumn(ws, "AfterSnapshotID") = 0 Then Exit Sub
    If HeaderColumn(snapshots, "EventID") = 0 Or HeaderColumn(snapshots, "EmployeeID") = 0 Or HeaderColumn(snapshots, "SnapshotKind") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        If RowHasAnyValue(ws, rowNum) Then
            eventID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EventID")).Value)
            employeeID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EmployeeID")).Value)
            beforeID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "BeforeSnapshotID")).Value)
            afterID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "AfterSnapshotID")).Value)
            CheckOneEventSnapshot eventID, employeeID, beforeID, "BEFORE", snapshotRows, snapshots, findings, rowNum
            CheckOneEventSnapshot eventID, employeeID, afterID, "AFTER", snapshotRows, snapshots, findings, rowNum
            If beforeID <> "" And beforeID = afterID Then AddFinding findings, "ERROR", "EVENT_SNAPSHOTS", "Event", RowToken(SHEET_EVENTS, rowNum), "Before and after snapshots are identical.", "Store two distinct snapshots for the event."
        End If
    Next rowNum
End Sub

Private Sub CheckOneEventSnapshot(ByVal eventID As String, ByVal employeeID As String, ByVal snapshotID As String, ByVal expectedKind As String, ByVal snapshotRows As Object, ByVal snapshots As Worksheet, ByVal findings As Collection, ByVal eventRow As Long)
    Dim snapshotRow As Long
    Dim actualEventID As String
    Dim actualEmployeeID As String
    Dim actualKind As String
    If snapshotID = "" Then
        AddFinding findings, "ERROR", "EVENT_SNAPSHOTS", "Event", RowToken(SHEET_EVENTS, eventRow), "Event snapshot reference is blank.", "Link the event to a valid before/after snapshot."
        Exit Sub
    End If
    If snapshotRows Is Nothing Or Not snapshotRows.Exists(snapshotID) Then
        AddFinding findings, "ERROR", "EVENT_SNAPSHOTS", "Event", RowToken(SHEET_EVENTS, eventRow), "Event references a missing snapshot.", "Restore the referenced snapshot or correct the event link."
        Exit Sub
    End If
    snapshotRow = CLng(snapshotRows(snapshotID))
    actualEventID = CellText(snapshots.Cells(snapshotRow, HeaderColumn(snapshots, "EventID")).Value)
    actualEmployeeID = CellText(snapshots.Cells(snapshotRow, HeaderColumn(snapshots, "EmployeeID")).Value)
    actualKind = UCase$(CellText(snapshots.Cells(snapshotRow, HeaderColumn(snapshots, "SnapshotKind")).Value))
    If actualEventID <> eventID Or actualEmployeeID <> employeeID Or actualKind <> expectedKind Then AddFinding findings, "ERROR", "EVENT_SNAPSHOTS", "Event", RowToken(SHEET_EVENTS, eventRow), "Event snapshot linkage does not match event, employee, or kind.", "Correct the snapshot linkage manually."
End Sub

Private Sub CheckSnapshotLinks(ByVal ws As Worksheet, ByVal eventRows As Object, ByVal employeeRows As Object, ByVal findings As Collection)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim eventID As String
    Dim employeeID As String
    Dim kindText As String
    If ws Is Nothing Or eventRows Is Nothing Or employeeRows Is Nothing Then Exit Sub
    If HeaderColumn(ws, "EventID") = 0 Or HeaderColumn(ws, "EmployeeID") = 0 Or HeaderColumn(ws, "SnapshotKind") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        If RowHasAnyValue(ws, rowNum) Then
            eventID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EventID")).Value)
            employeeID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EmployeeID")).Value)
            kindText = UCase$(CellText(ws.Cells(rowNum, HeaderColumn(ws, "SnapshotKind")).Value))
            If eventID = "" Or Not eventRows.Exists(eventID) Then AddFinding findings, "ERROR", "SNAPSHOT_LINKAGE", "Snapshot", RowToken(SHEET_SNAPSHOTS, rowNum), "Snapshot references a missing event.", "Link the snapshot to an existing event."
            If employeeID = "" Or Not employeeRows.Exists(employeeID) Then AddFinding findings, "ERROR", "SNAPSHOT_LINKAGE", "Snapshot", RowToken(SHEET_SNAPSHOTS, rowNum), "Snapshot references a missing employee.", "Link the snapshot to an existing employee."
            If kindText <> "BEFORE" And kindText <> "AFTER" Then AddFinding findings, "ERROR", "SNAPSHOT_LINKAGE", "Snapshot", RowToken(SHEET_SNAPSHOTS, rowNum), "Snapshot kind is not BEFORE or AFTER.", "Set the snapshot kind explicitly."
        End If
    Next rowNum
End Sub

Private Sub CheckAssignmentLinks(ByVal ws As Worksheet, ByVal employeeRows As Object, ByVal eventRows As Object, ByVal legalActRows As Object, ByVal findings As Collection)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeID As String
    Dim eventID As String
    Dim actID As String
    If ws Is Nothing Or employeeRows Is Nothing Or eventRows Is Nothing Or legalActRows Is Nothing Then Exit Sub
    If HeaderColumn(ws, "EmployeeID") = 0 Or HeaderColumn(ws, "EventID") = 0 Or HeaderColumn(ws, "ActID") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        If RowHasAnyValue(ws, rowNum) Then
            employeeID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EmployeeID")).Value)
            eventID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EventID")).Value)
            actID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "ActID")).Value)
            If employeeID = "" Or Not employeeRows.Exists(employeeID) Then AddFinding findings, "ERROR", "ASSIGNMENT_LINKAGE", "Assignment", RowToken(SHEET_ASSIGNMENTS, rowNum), "Payment assignment references a missing employee.", "Link the assignment to an existing employee."
            If eventID = "" Or Not eventRows.Exists(eventID) Then AddFinding findings, "ERROR", "ASSIGNMENT_LINKAGE", "Assignment", RowToken(SHEET_ASSIGNMENTS, rowNum), "Payment assignment references a missing event.", "Link the assignment to an existing event."
            If actID <> "" And (legalActRows Is Nothing Or Not legalActRows.Exists(actID)) Then AddFinding findings, "ERROR", "LEGAL_REFERENCE", "Assignment", RowToken(SHEET_ASSIGNMENTS, rowNum), "Payment assignment references a missing legal act.", "Register the legal act or correct the assignment reference."
        End If
    Next rowNum
End Sub

Private Sub CheckDocumentLinks(ByVal ws As Worksheet, ByVal eventRows As Object, ByVal findings As Collection)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim eventID As String
    Dim filePath As String
    If ws Is Nothing Or eventRows Is Nothing Then Exit Sub
    If HeaderColumn(ws, "EventID") = 0 Or HeaderColumn(ws, "FilePath") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        If RowHasAnyValue(ws, rowNum) Then
            eventID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EventID")).Value)
            filePath = CellText(ws.Cells(rowNum, HeaderColumn(ws, "FilePath")).Value)
            If eventID = "" Or Not eventRows.Exists(eventID) Then AddFinding findings, "ERROR", "DOCUMENT_LINKAGE", "Document", RowToken(SHEET_DOCUMENTS, rowNum), "Document references a missing event.", "Link the document to an existing event."
            If filePath = "" Or Not FileExistsSafe(filePath) Then AddFinding findings, "WARNING", "DOCUMENT_LINKAGE", "Document", RowToken(SHEET_DOCUMENTS, rowNum), "Registered document file is unavailable.", "Check the file manually and register a valid path if appropriate."
        End If
    Next rowNum
End Sub

Private Sub CheckCurrentStateLinks(ByVal ws As Worksheet, ByVal employeeRows As Object, ByVal eventRows As Object, ByVal findings As Collection)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeID As String
    Dim sourceEventID As String
    Dim lastEventID As String
    If ws Is Nothing Or employeeRows Is Nothing Or eventRows Is Nothing Then Exit Sub
    If HeaderColumn(ws, "EmployeeID") = 0 Or HeaderColumn(ws, "SourceEventID") = 0 Or HeaderColumn(ws, "LastEventID") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        If RowHasAnyValue(ws, rowNum) Then
            employeeID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EmployeeID")).Value)
            sourceEventID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "SourceEventID")).Value)
            lastEventID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "LastEventID")).Value)
            If employeeID = "" Or Not employeeRows.Exists(employeeID) Then AddFinding findings, "ERROR", "CURRENT_STATE_LINKAGE", "CurrentState", RowToken(SHEET_CURRENT_STATE, rowNum), "Current state references a missing employee.", "Link the current state to an existing employee."
            If lastEventID = "" Or Not eventRows.Exists(lastEventID) Then
                AddFinding findings, "ERROR", "CURRENT_STATE_LINKAGE", "CurrentState", RowToken(SHEET_CURRENT_STATE, rowNum), "Current state references a missing LastEventID.", "Link LastEventID to the event that produced this state."
            End If
            If sourceEventID <> "" And (eventRows Is Nothing Or Not eventRows.Exists(sourceEventID)) Then AddFinding findings, "ERROR", "CURRENT_STATE_LINKAGE", "CurrentState", RowToken(SHEET_CURRENT_STATE, rowNum), "Current state references a missing SourceEventID.", "Correct the source event reference."
        End If
    Next rowNum
End Sub

Private Sub CheckChronology(ByVal events As Worksheet, ByVal eventRows As Object, ByVal snapshots As Worksheet, ByVal findings As Collection)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim eventDate As Variant
    Dim effectiveDate As Variant
    Dim eventCreated As Variant
    Dim snapshotCreated As Variant
    Dim eventID As String
    Dim snapshotEventID As String
    If events Is Nothing Or snapshots Is Nothing Then Exit Sub
    If HeaderColumn(events, "EventID") = 0 Or HeaderColumn(events, "EventDate") = 0 Or HeaderColumn(events, "EffectiveDate") = 0 Or HeaderColumn(events, "CreatedAt") = 0 Or HeaderColumn(events, "BeforeSnapshotID") = 0 Or HeaderColumn(events, "AfterSnapshotID") = 0 Then Exit Sub
    If HeaderColumn(snapshots, "SnapshotID") = 0 Or HeaderColumn(snapshots, "CreatedAt") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(events)
    For rowNum = 2 To lastRow
        If RowHasAnyValue(events, rowNum) Then
            eventID = CellText(events.Cells(rowNum, HeaderColumn(events, "EventID")).Value)
            eventDate = events.Cells(rowNum, HeaderColumn(events, "EventDate")).Value
            effectiveDate = events.Cells(rowNum, HeaderColumn(events, "EffectiveDate")).Value
            eventCreated = events.Cells(rowNum, HeaderColumn(events, "CreatedAt")).Value
            If IsDate(eventDate) And IsDate(effectiveDate) Then
                If CDate(effectiveDate) < CDate(eventDate) Then AddFinding findings, "ERROR", "CHRONOLOGY", "Event", RowToken(SHEET_EVENTS, rowNum), "Effective date precedes event date.", "Correct the event chronology before exporting documents."
            End If
            If Not snapshots Is Nothing Then
                snapshotCreated = FindSnapshotCreatedAt(snapshots, CellText(events.Cells(rowNum, HeaderColumn(events, "BeforeSnapshotID")).Value))
                If IsDate(eventCreated) And IsDate(snapshotCreated) Then
                    If CDate(snapshotCreated) < CDate(eventCreated) Then AddFinding findings, "ERROR", "CHRONOLOGY", "Event", RowToken(SHEET_EVENTS, rowNum), "Snapshot was created before its event record.", "Review the event and snapshot timestamps."
                End If
                snapshotCreated = FindSnapshotCreatedAt(snapshots, CellText(events.Cells(rowNum, HeaderColumn(events, "AfterSnapshotID")).Value))
                If IsDate(eventCreated) And IsDate(snapshotCreated) Then
                    If CDate(snapshotCreated) < CDate(eventCreated) Then AddFinding findings, "ERROR", "CHRONOLOGY", "Event", RowToken(SHEET_EVENTS, rowNum), "Snapshot was created before its event record.", "Review the event and snapshot timestamps."
                End If
            End If
        End If
    Next rowNum
End Sub

Private Function BuildExcludedEmployeeIndex(ByVal ws As Worksheet) As Object
    Dim index As Object
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeID As String
    Dim eventType As String
    Dim statusText As String
    Set index = CreateObject("Scripting.Dictionary")
    index.CompareMode = vbTextCompare
    If ws Is Nothing Then
        Set BuildExcludedEmployeeIndex = index
        Exit Function
    End If
    lastRow = IntegrityLastUsedRow(ws)
    For rowNum = 2 To lastRow
        eventType = UCase$(CellText(ws.Cells(rowNum, HeaderColumn(ws, "EventType")).Value))
        statusText = UCase$(CellText(ws.Cells(rowNum, HeaderColumn(ws, "Status")).Value))
        employeeID = CellText(ws.Cells(rowNum, HeaderColumn(ws, "EmployeeID")).Value)
        If eventType = "EXCLUSION" And statusText <> "CANCELLED" And employeeID <> "" Then
            If Not index.Exists(employeeID) Then index.Add employeeID, True
        End If
    Next rowNum
    Set BuildExcludedEmployeeIndex = index
End Function

Private Sub CheckExclusionAndActivePayments(ByVal employees As Worksheet, ByVal assignments As Worksheet, ByVal excludedEmployees As Object, ByVal findings As Collection)
    Dim employeeStatus As Object
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeID As String
    Dim statusText As String
    Dim assignmentStatus As String
    Set employeeStatus = CreateObject("Scripting.Dictionary")
    employeeStatus.CompareMode = vbTextCompare
    If Not employees Is Nothing And HeaderColumn(employees, "EmployeeID") > 0 And HeaderColumn(employees, "IsActive") > 0 Then
        lastRow = IntegrityLastUsedRow(employees)
        For rowNum = 2 To lastRow
            employeeID = CellText(employees.Cells(rowNum, HeaderColumn(employees, "EmployeeID")).Value)
            If employeeID <> "" Then employeeStatus(employeeID) = UCase$(CellText(employees.Cells(rowNum, HeaderColumn(employees, "IsActive")).Value))
        Next rowNum
    End If
    If Not assignments Is Nothing And HeaderColumn(assignments, "EmployeeID") > 0 And HeaderColumn(assignments, "Status") > 0 Then
        lastRow = IntegrityLastUsedRow(assignments)
        For rowNum = 2 To lastRow
            employeeID = CellText(assignments.Cells(rowNum, HeaderColumn(assignments, "EmployeeID")).Value)
            assignmentStatus = UCase$(CellText(assignments.Cells(rowNum, HeaderColumn(assignments, "Status")).Value))
            If assignmentStatus = "ACTIVE" And ((employeeStatus.Exists(employeeID) And employeeStatus(employeeID) = "NO") Or (Not excludedEmployees Is Nothing And excludedEmployees.Exists(employeeID))) Then
                AddFinding findings, "ERROR", "ACTIVE_PAYMENTS", "Assignment", RowToken(SHEET_ASSIGNMENTS, rowNum), "Active payment remains for an excluded or inactive employee.", "Terminate or correct the assignment after reviewing the exclusion event."
            End If
        Next rowNum
    End If
    If Not employees Is Nothing And Not excludedEmployees Is Nothing And HeaderColumn(employees, "EmployeeID") > 0 And HeaderColumn(employees, "IsActive") > 0 Then
        For rowNum = 2 To IntegrityLastUsedRow(employees)
            employeeID = CellText(employees.Cells(rowNum, HeaderColumn(employees, "EmployeeID")).Value)
            statusText = UCase$(CellText(employees.Cells(rowNum, HeaderColumn(employees, "IsActive")).Value))
            If employeeID <> "" And excludedEmployees.Exists(employeeID) And statusText <> "NO" Then AddFinding findings, "ERROR", "EXCLUSION_STATE", "Employee", RowToken(SHEET_EMPLOYEES, rowNum), "Excluded employee is still marked active.", "Set the employee state consistently with the exclusion event."
        Next rowNum
    End If
End Sub

Private Sub CheckStaffLinks(ByVal employees As Worksheet, ByVal findings As Collection)
    Dim staff As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim employeeID As String
    Dim linkStatus As String
    Dim referenceText As String
    Dim staffRow As Long
    Set staff = GetIntegritySheet(StaffSheetName())
    If employees Is Nothing Or staff Is Nothing Then Exit Sub
    If HeaderColumn(employees, "EmployeeID") = 0 Or HeaderColumn(employees, "StaffLinkStatus") = 0 Or HeaderColumn(employees, "StaffReference") = 0 Then Exit Sub
    lastRow = IntegrityLastUsedRow(employees)
    For rowNum = 2 To lastRow
        employeeID = CellText(employees.Cells(rowNum, HeaderColumn(employees, "EmployeeID")).Value)
        linkStatus = UCase$(CellText(employees.Cells(rowNum, HeaderColumn(employees, "StaffLinkStatus")).Value))
        referenceText = CellText(employees.Cells(rowNum, HeaderColumn(employees, "StaffReference")).Value)
        If linkStatus = "LINKED" Or linkStatus = "CONFIRMED" Then
            staffRow = ParseStaffRow(referenceText)
            If staffRow < 2 Or staffRow > IntegrityLastUsedRow(staff) Then AddFinding findings, "WARNING", "STAFF_LINK", "Employee", RowToken(SHEET_EMPLOYEES, rowNum), "Confirmed staff link points outside the staff sheet.", "Review the manual staff link before synchronizing state."
        End If
    Next rowNum
End Sub

Private Sub CheckLegalAndClassificationReferences(ByVal assignments As Worksheet, ByVal legalActs As Worksheet, ByVal legalActRows As Object, ByVal findings As Collection)
    Dim rules As Worksheet
    Dim classifications As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim actID As String
    Dim reviewStatus As String
    Set rules = GetIntegritySheet(SHEET_PAYMENT_RULES)
    If Not rules Is Nothing And Not legalActRows Is Nothing And HeaderColumn(rules, "ActID") > 0 Then
        lastRow = IntegrityLastUsedRow(rules)
        For rowNum = 2 To lastRow
            actID = CellText(rules.Cells(rowNum, HeaderColumn(rules, "ActID")).Value)
            If actID <> "" And Not legalActRows.Exists(actID) Then AddFinding findings, "ERROR", "LEGAL_REFERENCE", "PaymentRule", RowToken(SHEET_PAYMENT_RULES, rowNum), "Payment rule references a missing legal act.", "Register the legal act or keep the rule inactive."
        Next rowNum
    End If
    Set classifications = GetIntegritySheet(SHEET_POSITION_CLASSIFICATION)
    If Not classifications Is Nothing And HeaderColumn(classifications, "ReviewStatus") > 0 And HeaderColumn(classifications, "PositionKey") > 0 Then
        lastRow = IntegrityLastUsedRow(classifications)
        For rowNum = 2 To lastRow
            reviewStatus = UCase$(CellText(classifications.Cells(rowNum, HeaderColumn(classifications, "ReviewStatus")).Value))
            If reviewStatus = "ACTIVE" And CellText(classifications.Cells(rowNum, HeaderColumn(classifications, "PositionKey")).Value) = "" Then AddFinding findings, "WARNING", "LEGAL_REFERENCE", "Classification", RowToken(SHEET_POSITION_CLASSIFICATION, rowNum), "Active classification has no position key.", "Review the classification before using it in calculations."
        Next rowNum
    End If
End Sub

Private Function FindSnapshotCreatedAt(ByVal ws As Worksheet, ByVal snapshotID As String) As Variant
    Dim rowNum As Long
    If ws Is Nothing Or snapshotID = "" Or HeaderColumn(ws, "SnapshotID") = 0 Or HeaderColumn(ws, "CreatedAt") = 0 Then Exit Function
    For rowNum = 2 To IntegrityLastUsedRow(ws)
        If CellText(ws.Cells(rowNum, HeaderColumn(ws, "SnapshotID")).Value) = snapshotID Then
            FindSnapshotCreatedAt = ws.Cells(rowNum, HeaderColumn(ws, "CreatedAt")).Value
            Exit Function
        End If
    Next rowNum
End Function

Private Function GetIntegritySheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetIntegritySheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function HeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim colNum As Long
    If ws Is Nothing Then Exit Function
    For colNum = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If StrComp(CellText(ws.Cells(1, colNum).Value), headerName, vbTextCompare) = 0 Then
            HeaderColumn = colNum
            Exit Function
        End If
    Next colNum
End Function

Private Function IntegrityLastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If foundCell Is Nothing Then
        IntegrityLastUsedRow = 1
    Else
        IntegrityLastUsedRow = foundCell.Row
    End If
End Function

Private Function IntegrityDataRowCount(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = IntegrityLastUsedRow(ws)
    If lastRow >= 2 Then IntegrityDataRowCount = lastRow - 1
End Function

Private Function RowHasAnyValue(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim lastColumn As Long
    Dim colNum As Long
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For colNum = 1 To lastColumn
        If CellText(ws.Cells(rowNum, colNum).Value) <> "" Then
            RowHasAnyValue = True
            Exit Function
        End If
    Next colNum
End Function

Private Function FileExistsSafe(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExistsSafe = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    On Error GoTo 0
End Function

Private Function ParseStaffRow(ByVal referenceText As String) As Long
    If Left$(referenceText, Len(STAFF_ROW_PREFIX)) <> STAFF_ROW_PREFIX Then Exit Function
    If IsNumeric(Mid$(referenceText, Len(STAFF_ROW_PREFIX) + 1)) Then ParseStaffRow = CLng(Mid$(referenceText, Len(STAFF_ROW_PREFIX) + 1))
End Function

Private Function StaffSheetName() As String
    StaffSheetName = ChrW$(1064) & ChrW$(1090) & ChrW$(1072) & ChrW$(1090)
End Function

Private Function RowToken(ByVal sheetName As String, ByVal rowNum As Long) As String
    RowToken = sheetName & "!row " & CStr(rowNum)
End Function

Private Function CellText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    CellText = Trim$(CStr(rawValue))
End Function

Private Function SafeIntegrityText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    SafeIntegrityText = CStr(rawValue)
End Function

Private Function ValueOrZero(ByVal source As Object, ByVal key As String) As Long
    If source Is Nothing Then Exit Function
    If source.Exists(key) Then ValueOrZero = CLng(source(key))
End Function

Private Sub AddFinding(ByVal findings As Collection, ByVal severity As String, ByVal category As String, ByVal entityType As String, ByVal entityID As String, ByVal messageText As String, ByVal suggestedAction As String)
    Dim finding As Object
    Set finding = CreateObject("Scripting.Dictionary")
    finding.CompareMode = vbTextCompare
    finding.Add "Severity", severity
    finding.Add "Category", category
    finding.Add "EntityType", entityType
    finding.Add "EntityID", entityID
    finding.Add "Message", messageText
    finding.Add "SuggestedAction", suggestedAction
    findings.Add finding
    If UCase$(severity) = "WARNING" Or UCase$(severity) = "ERROR" Then Debug.Print "WARN integrity-finding|severity=" & severity & "|category=" & category & "|entity=" & entityType
End Sub

Private Sub SortFindings(ByVal findings As Collection)
    Dim items() As Object
    Dim i As Long
    Dim j As Long
    Dim temp As Object
    If findings Is Nothing Or findings.Count < 2 Then Exit Sub
    ReDim items(1 To findings.Count)
    For i = 1 To findings.Count
        Set items(i) = findings(i)
    Next i
    For i = 1 To UBound(items) - 1
        For j = i + 1 To UBound(items)
            If FindingSortKey(items(j)) < FindingSortKey(items(i)) Then
                Set temp = items(i)
                Set items(i) = items(j)
                Set items(j) = temp
            End If
        Next j
    Next i
    Do While findings.Count > 0
        findings.Remove 1
    Loop
    For i = 1 To UBound(items)
        findings.Add items(i)
    Next i
End Sub

Private Function FindingSortKey(ByVal finding As Object) As String
    Dim rank As String
    Select Case UCase$(SafeIntegrityText(finding("Severity")))
        Case "ERROR": rank = "1"
        Case "WARNING": rank = "2"
        Case Else: rank = "3"
    End Select
    FindingSortKey = rank & "|" & UCase$(SafeIntegrityText(finding("Category"))) & "|" & UCase$(SafeIntegrityText(finding("EntityType"))) & "|" & UCase$(SafeIntegrityText(finding("EntityID")))
End Function

Private Sub LogIntegrityComplete(ByVal findings As Collection)
    Dim item As Object
    Dim errors As Long
    Dim warnings As Long
    For Each item In findings
        If UCase$(SafeIntegrityText(item("Severity"))) = "ERROR" Then errors = errors + 1
        If UCase$(SafeIntegrityText(item("Severity"))) = "WARNING" Then warnings = warnings + 1
    Next item
    Debug.Print "INFO integrity-scan-complete|findings=" & CStr(findings.Count) & "|errors=" & CStr(errors) & "|warnings=" & CStr(warnings)
End Sub
