Attribute VB_Name = "mdlLegalActs"
Option Explicit

' Legal-act registry input. Existing entries remain immutable; a new revision
' is stored as a new row with its own ActID.

Private Const INPUT_SHEET As String = "LegalActInput"
Private Const ACTS_SHEET As String = "LegalActs"

Public Sub OpenLegalActInput()
    EnsureLegalActInputSheet
    ThisWorkbook.Worksheets(INPUT_SHEET).Activate
End Sub

Public Function SaveLegalActInput(Optional ByVal showMessage As Boolean = True) As String
    Dim inputSheet As Worksheet
    Dim actsSheet As Worksheet
    Dim actID As String
    Dim rowNum As Long

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    EnsureLegalActInputSheet
    Set inputSheet = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set actsSheet = ThisWorkbook.Worksheets(ACTS_SHEET)

    actID = ActText(GetInputValue(inputSheet, "act_id"))
    If actID = "" Then actID = BuildActID()
    If ActExists(actsSheet, actID) Then Err.Raise vbObjectError + 780, "mdlLegalActs", "ActID already exists. Save a new revision with a new ActID."
    ValidateInput inputSheet

    rowNum = actsSheet.Cells(actsSheet.Rows.Count, 1).End(xlUp).Row + 1
    If rowNum < 2 Then rowNum = 2
    actsSheet.Cells(rowNum, 1).Value = actID
    actsSheet.Cells(rowNum, 2).Value = GetInputValue(inputSheet, "act_type")
    actsSheet.Cells(rowNum, 3).Value = GetInputValue(inputSheet, "act_number")
    actsSheet.Cells(rowNum, 4).Value = GetInputValue(inputSheet, "act_date")
    actsSheet.Cells(rowNum, 5).Value = GetInputValue(inputSheet, "title")
    actsSheet.Cells(rowNum, 6).Value = GetInputValue(inputSheet, "revision")
    actsSheet.Cells(rowNum, 7).Value = GetInputValue(inputSheet, "effective_from")
    actsSheet.Cells(rowNum, 8).Value = GetInputValue(inputSheet, "effective_to")
    actsSheet.Cells(rowNum, 9).Value = GetInputValue(inputSheet, "access_mark")
    actsSheet.Cells(rowNum, 10).Value = GetInputValue(inputSheet, "note")
    actsSheet.Cells(rowNum, 11).Value = Now
    actsSheet.Cells(rowNum, 12).Value = Now

    SetInputValue inputSheet, "act_id", actID
    SetInputValue inputSheet, "saved_act_id", actID
    If showMessage Then MsgBox "Legal act was saved: " & actID, vbInformation
    SaveLegalActInput = actID
End Function

Private Sub EnsureLegalActInputSheet()
    Dim ws As Worksheet
    Dim fields As Variant
    Dim labels As Variant
    Dim index As Long

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = INPUT_SHEET
    End If

    fields = Array("act_id", "act_type", "act_number", "act_date", "title", "revision", "effective_from", "effective_to", "access_mark", "note", "saved_act_id")
    labels = Array("Act ID (optional for new act)", "Document type", "Number", "Date", "Title", "Revision", "Effective from", "Effective to", "Access mark", "Note", "Saved Act ID")
    If ActText(ws.Cells(1, 1).Value) = "" Then
        ws.Cells(1, 1).Value = "LEGAL ACT INPUT"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(2, 1).Value = "This form creates a new legal-act record only. Use a new Act ID for a new revision."
    End If
    ws.Cells(3, 1).Value = "Field"
    ws.Cells(3, 2).Value = "Value"
    ws.Rows(3).Font.Bold = True
    ws.Rows(3).Interior.Color = RGB(217, 225, 242)
    For index = LBound(fields) To UBound(fields)
        ws.Cells(index + 4, 1).Value = fields(index)
        If ActText(ws.Cells(index + 4, 3).Value) = "" Then ws.Cells(index + 4, 3).Value = labels(index)
    Next index
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 32
    ws.Columns(3).ColumnWidth = 44
End Sub

Private Sub ValidateInput(ByVal ws As Worksheet)
    If ActText(GetInputValue(ws, "act_type")) = "" Then Err.Raise vbObjectError + 781, "mdlLegalActs", "Document type is required."
    If ActText(GetInputValue(ws, "title")) = "" Then Err.Raise vbObjectError + 782, "mdlLegalActs", "Title is required."
    If Not IsDate(GetInputValue(ws, "act_date")) Then Err.Raise vbObjectError + 783, "mdlLegalActs", "Act date is required and must be a date."
    If ActText(GetInputValue(ws, "effective_from")) <> "" And Not IsDate(GetInputValue(ws, "effective_from")) Then Err.Raise vbObjectError + 784, "mdlLegalActs", "Effective-from value must be a date."
    If ActText(GetInputValue(ws, "effective_to")) <> "" And Not IsDate(GetInputValue(ws, "effective_to")) Then Err.Raise vbObjectError + 785, "mdlLegalActs", "Effective-to value must be a date."
    If IsDate(GetInputValue(ws, "effective_from")) And IsDate(GetInputValue(ws, "effective_to")) Then
        If CDate(GetInputValue(ws, "effective_to")) < CDate(GetInputValue(ws, "effective_from")) Then Err.Raise vbObjectError + 786, "mdlLegalActs", "Effective-to cannot be earlier than effective-from."
    End If
End Sub

Private Function ActExists(ByVal ws As Worksheet, ByVal actID As String) As Boolean
    Dim rowNum As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(ActText(ws.Cells(rowNum, 1).Value), actID, vbTextCompare) = 0 Then
            ActExists = True
            Exit Function
        End If
    Next rowNum
End Function

Private Function BuildActID() As String
    BuildActID = "ACT-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(Int((Timer * 100) Mod 100))
End Function

Private Function FindInputRow(ByVal ws As Worksheet, ByVal fieldName As String) As Long
    Dim rowNum As Long
    For rowNum = 4 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(ActText(ws.Cells(rowNum, 1).Value), fieldName, vbTextCompare) = 0 Then
            FindInputRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Function GetInputValue(ByVal ws As Worksheet, ByVal fieldName As String) As Variant
    Dim rowNum As Long
    rowNum = FindInputRow(ws, fieldName)
    If rowNum > 0 Then GetInputValue = ws.Cells(rowNum, 2).Value
End Function

Private Sub SetInputValue(ByVal ws As Worksheet, ByVal fieldName As String, ByVal fieldValue As Variant)
    Dim rowNum As Long
    rowNum = FindInputRow(ws, fieldName)
    If rowNum > 0 Then ws.Cells(rowNum, 2).Value = fieldValue
End Sub

Private Function ActText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    ActText = Trim$(CStr(value))
End Function
