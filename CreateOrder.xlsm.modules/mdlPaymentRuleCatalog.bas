Attribute VB_Name = "mdlPaymentRuleCatalog"
Option Explicit

' Catalog input only. Saved rules stay DRAFT and are not evaluated until a
' separately approved limited-operator implementation is added.

Private Const INPUT_SHEET As String = "PaymentRuleInput"
Private Const RULES_SHEET As String = "PaymentRules"

Public Sub OpenPaymentRuleInput()
    EnsurePaymentRuleInputSheet
    ThisWorkbook.Worksheets(INPUT_SHEET).Activate
End Sub

Public Function SavePaymentRuleInput(Optional ByVal showMessage As Boolean = True) As String
    Dim inputSheet As Worksheet
    Dim rulesSheet As Worksheet
    Dim ruleID As String
    Dim rowNum As Long
    Dim fields As Variant
    Dim index As Long

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    EnsurePaymentRuleInputSheet
    Set inputSheet = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set rulesSheet = ThisWorkbook.Worksheets(RULES_SHEET)
    ruleID = RuleText(GetInputValue(inputSheet, "rule_id"))
    If ruleID = "" Then ruleID = BuildRuleID()
    If RuleExists(rulesSheet, ruleID) Then Err.Raise vbObjectError + 800, "mdlPaymentRuleCatalog", "RuleID already exists. Create a new version with a new RuleID."
    ValidateInput inputSheet

    rowNum = rulesSheet.Cells(rulesSheet.Rows.Count, 1).End(xlUp).Row + 1
    If rowNum < 2 Then rowNum = 2
    fields = Array("payment_code", "basis_code", "amount_kind", "amount_value", "condition_type", "condition_operator", "expected_value", "fact_source", "required_documents", "start_date_source", "end_rule", "act_id", "act_point", "cap_group", "priority", "severity", "explanation_template", "word_template", "effective_from", "effective_to")
    rulesSheet.Cells(rowNum, 1).Value = ruleID
    For index = LBound(fields) To UBound(fields)
        rulesSheet.Cells(rowNum, index + 2).Value = GetInputValue(inputSheet, fields(index))
    Next index
    rulesSheet.Cells(rowNum, 22).Value = "DRAFT"
    rulesSheet.Cells(rowNum, 23).Value = Now
    rulesSheet.Cells(rowNum, 24).Value = Now
    SetInputValue inputSheet, "rule_id", ruleID
    SetInputValue inputSheet, "saved_rule_id", ruleID
    If showMessage Then MsgBox "Payment rule was saved as DRAFT: " & ruleID, vbInformation
    SavePaymentRuleInput = ruleID
End Function

Private Sub EnsurePaymentRuleInputSheet()
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
    fields = Array("rule_id", "payment_code", "basis_code", "amount_kind", "amount_value", "condition_type", "condition_operator", "expected_value", "fact_source", "required_documents", "start_date_source", "end_rule", "act_id", "act_point", "cap_group", "priority", "severity", "explanation_template", "word_template", "effective_from", "effective_to", "saved_rule_id")
    labels = Array("Rule ID (optional for new version)", "Payment code", "Basis code", "Amount kind", "Amount value", "Condition type", "Condition operator", "Expected value", "Fact source", "Required documents", "Start date source", "End rule", "Act ID", "Act point", "Cap group", "Priority", "Severity", "Explanation template", "Word template", "Effective from", "Effective to", "Saved Rule ID")
    If RuleText(ws.Cells(1, 1).Value) = "" Then
        ws.Cells(1, 1).Value = "PAYMENT RULE INPUT"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(2, 1).Value = "Saving creates a DRAFT catalog row only. It does not activate a payment calculation."
    End If
    ws.Cells(3, 1).Value = "Field"
    ws.Cells(3, 2).Value = "Value"
    ws.Rows(3).Font.Bold = True
    ws.Rows(3).Interior.Color = RGB(217, 225, 242)
    For index = LBound(fields) To UBound(fields)
        ws.Cells(index + 4, 1).Value = fields(index)
        If RuleText(ws.Cells(index + 4, 3).Value) = "" Then ws.Cells(index + 4, 3).Value = labels(index)
    Next index
    ws.Columns(1).ColumnWidth = 24
    ws.Columns(2).ColumnWidth = 32
    ws.Columns(3).ColumnWidth = 44
End Sub

Private Sub ValidateInput(ByVal ws As Worksheet)
    If RuleText(GetInputValue(ws, "payment_code")) = "" Then Err.Raise vbObjectError + 801, "mdlPaymentRuleCatalog", "Payment code is required."
    If RuleText(GetInputValue(ws, "basis_code")) = "" Then Err.Raise vbObjectError + 802, "mdlPaymentRuleCatalog", "Basis code is required."
    If RuleText(GetInputValue(ws, "amount_kind")) = "" Then Err.Raise vbObjectError + 803, "mdlPaymentRuleCatalog", "Amount kind is required."
    If RuleText(GetInputValue(ws, "effective_from")) <> "" And Not IsDate(GetInputValue(ws, "effective_from")) Then Err.Raise vbObjectError + 804, "mdlPaymentRuleCatalog", "Effective-from value must be a date."
    If RuleText(GetInputValue(ws, "effective_to")) <> "" And Not IsDate(GetInputValue(ws, "effective_to")) Then Err.Raise vbObjectError + 805, "mdlPaymentRuleCatalog", "Effective-to value must be a date."
    If IsDate(GetInputValue(ws, "effective_from")) And IsDate(GetInputValue(ws, "effective_to")) Then
        If CDate(GetInputValue(ws, "effective_to")) < CDate(GetInputValue(ws, "effective_from")) Then Err.Raise vbObjectError + 806, "mdlPaymentRuleCatalog", "Effective-to cannot be earlier than effective-from."
    End If
End Sub

Private Function RuleExists(ByVal ws As Worksheet, ByVal ruleID As String) As Boolean
    Dim rowNum As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If StrComp(RuleText(ws.Cells(rowNum, 1).Value), ruleID, vbTextCompare) = 0 Then RuleExists = True: Exit Function
    Next rowNum
End Function

Private Function BuildRuleID() As String
    BuildRuleID = "RULE-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(Int((Timer * 100) Mod 100))
End Function

Private Function FindInputRow(ByVal ws As Worksheet, ByVal fieldName As String) As Long
    Dim rowNum As Long
    For rowNum = 4 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(RuleText(ws.Cells(rowNum, 1).Value), fieldName, vbTextCompare) = 0 Then FindInputRow = rowNum: Exit Function
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

Private Function RuleText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    RuleText = Trim$(CStr(value))
End Function
