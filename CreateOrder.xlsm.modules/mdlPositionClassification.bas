Attribute VB_Name = "mdlPositionClassification"
Option Explicit

Private Const INPUT_SHEET As String = "PositionClassificationInput"
Private Const DATA_SHEET As String = "PositionClassification"

Public Sub OpenPositionClassificationInput()
    EnsureInputSheet
    ThisWorkbook.Worksheets(INPUT_SHEET).Activate
End Sub

Public Function SavePositionClassificationInput(Optional ByVal showMessage As Boolean = True) As String
    Dim inputSheet As Worksheet, dataSheet As Worksheet
    Dim recordID As String, rowNum As Long, fields As Variant, index As Long

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    EnsureInputSheet
    Set inputSheet = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set dataSheet = ThisWorkbook.Worksheets(DATA_SHEET)
    recordID = TextOf(ValueFor(inputSheet, "classification_id"))
    If recordID = "" Then recordID = "POS-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(Int((Timer * 100) Mod 100))
    If ExistsID(dataSheet, recordID) Then Err.Raise vbObjectError + 820, "mdlPositionClassification", "ClassificationID already exists."
    If TextOf(ValueFor(inputSheet, "position_key")) = "" And TextOf(ValueFor(inputSheet, "staff_code")) = "" Then Err.Raise vbObjectError + 821, "mdlPositionClassification", "Position key or staff code is required."

    rowNum = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row + 1
    If rowNum < 2 Then rowNum = 2
    dataSheet.Cells(rowNum, 1).Value = recordID
    fields = Array("position_key", "staff_code", "position_text", "group_code", "command_level", "other_flags", "source", "review_status", "note")
    For index = LBound(fields) To UBound(fields)
        dataSheet.Cells(rowNum, index + 2).Value = ValueFor(inputSheet, fields(index))
    Next index
    If TextOf(dataSheet.Cells(rowNum, 9).Value) = "" Then dataSheet.Cells(rowNum, 9).Value = "DRAFT"
    dataSheet.Cells(rowNum, 11).Value = Now
    dataSheet.Cells(rowNum, 12).Value = Now
    SetValue inputSheet, "classification_id", recordID
    SetValue inputSheet, "saved_classification_id", recordID
    If showMessage Then MsgBox "Position classification was saved: " & recordID, vbInformation
    SavePositionClassificationInput = recordID
End Function

Private Sub EnsureInputSheet()
    Dim ws As Worksheet, fields As Variant, labels As Variant, index As Long
    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1)): ws.Name = INPUT_SHEET
    fields = Array("classification_id", "position_key", "staff_code", "position_text", "group_code", "command_level", "other_flags", "source", "review_status", "note", "saved_classification_id")
    labels = Array("Classification ID", "Normalized position key", "Staff code", "Position text", "Group code", "Command level", "Other flags", "Source", "Review status", "Note", "Saved Classification ID")
    If TextOf(ws.Cells(1, 1).Value) = "" Then ws.Cells(1, 1).Value = "POSITION CLASSIFICATION INPUT": ws.Cells(1, 1).Font.Bold = True: ws.Cells(2, 1).Value = "This catalog does not assign payments automatically."
    ws.Cells(3, 1).Value = "Field": ws.Cells(3, 2).Value = "Value": ws.Rows(3).Font.Bold = True
    For index = LBound(fields) To UBound(fields)
        ws.Cells(index + 4, 1).Value = fields(index)
        If TextOf(ws.Cells(index + 4, 3).Value) = "" Then ws.Cells(index + 4, 3).Value = labels(index)
    Next index
    ws.Columns(1).ColumnWidth = 26: ws.Columns(2).ColumnWidth = 32: ws.Columns(3).ColumnWidth = 40
End Sub

Private Function ExistsID(ByVal ws As Worksheet, ByVal recordID As String) As Boolean
    Dim rowNum As Long
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(TextOf(ws.Cells(rowNum, 1).Value), recordID, vbTextCompare) = 0 Then ExistsID = True: Exit Function
    Next rowNum
End Function

Private Function FieldRow(ByVal ws As Worksheet, ByVal fieldName As String) As Long
    Dim rowNum As Long
    For rowNum = 4 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(TextOf(ws.Cells(rowNum, 1).Value), fieldName, vbTextCompare) = 0 Then FieldRow = rowNum: Exit Function
    Next rowNum
End Function

Private Function ValueFor(ByVal ws As Worksheet, ByVal fieldName As String) As Variant
    Dim rowNum As Long: rowNum = FieldRow(ws, fieldName)
    If rowNum > 0 Then ValueFor = ws.Cells(rowNum, 2).Value
End Function

Private Sub SetValue(ByVal ws As Worksheet, ByVal fieldName As String, ByVal fieldValue As Variant)
    Dim rowNum As Long: rowNum = FieldRow(ws, fieldName)
    If rowNum > 0 Then ws.Cells(rowNum, 2).Value = fieldValue
End Sub

Private Function TextOf(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    TextOf = Trim$(CStr(value))
End Function
