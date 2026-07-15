Attribute VB_Name = "mdlPersonnelAllowanceRules"
Option Explicit

Public Const ALLOWANCE_STATUS_ACTIVE As String = "ACTIVE"
Public Const ALLOWANCE_STATUS_NOT_APPLICABLE As String = "NOT_APPLICABLE"
Public Const ALLOWANCE_STATUS_REQUIRES_DECISION As String = "REQUIRES_DECISION"
Public Const ALLOWANCE_STATUS_PENDING_LEGAL_ACT As String = "PENDING_LEGAL_ACT"

Public Function EvaluatePersonnelAllowances(ByVal stateData As Object, ByVal ruleData As Object) As Collection
    Dim results As New Collection
    Dim serviceCategory As String
    Dim point2Total As Double

    serviceCategory = UCase$(TextValue(stateData, "service_category"))
    AddFizoResult results, stateData, ruleData, serviceCategory
    AddMedalResult results, stateData, ruleData
    AddVusResult results, stateData, ruleData
    AddDriverResult results, stateData, ruleData
    AddTariffResult results, stateData, ruleData
    AddContract430Result results, stateData, ruleData
    ApplyPoint2Cap results, point2Total
    AddMobilizedFixedPayment results, stateData, ruleData, serviceCategory
    Set EvaluatePersonnelAllowances = results
End Function

Private Sub AddFizoResult(ByVal results As Collection, ByVal stateData As Object, ByVal ruleData As Object, ByVal serviceCategory As String)
    Dim levelCode As String
    Dim sportCode As String
    Dim amountValue As Double
    Dim statusValue As String
    Dim explanation As String

    If serviceCategory = "MOBILIZED" Then
        results.Add NewAllowance("FIZO", "FIZO", "PERCENT", 0, "SPECIAL_ACHIEVEMENTS_P2", ALLOWANCE_STATUS_NOT_APPLICABLE, "FIZO is excluded for MOBILIZED category.", ruleData)
        SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 2 Правил"
        Exit Sub
    End If

    levelCode = UCase$(TextValue(stateData, "fizo_level"))
    sportCode = UCase$(TextValue(stateData, "sport_status"))
    Select Case levelCode
        Case "SECOND": amountValue = 80
        Case "FIRST": amountValue = 90
        Case "HIGH": amountValue = 100
        Case Else: Exit Sub
    End Select

    statusValue = ALLOWANCE_STATUS_ACTIVE
    explanation = "FIZO rule resolved from qualification level and sport status."
    results.Add NewAllowance("FIZO", "FIZO_" & levelCode, "PERCENT", amountValue, "SPECIAL_ACHIEVEMENTS_P2", statusValue, explanation, ruleData)
    SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 2 Правил"
End Sub

Private Sub AddMedalResult(ByVal results As Collection, ByVal stateData As Object, ByVal ruleData As Object)
    Dim medalCode As String
    Dim amountValue As Double
    Dim paymentCode As String
    Dim awardDate As Variant
    Dim allowance As Object

    medalCode = UCase$(TextValue(stateData, "medal_code"))
    Select Case medalCode
        Case "COMBAT_DISTINCTION": paymentCode = "MEDAL_COMBAT_DISTINCTION": amountValue = 30
        Case "DEMINING": paymentCode = "MEDAL_DEMINING": amountValue = 20
        Case "MILITARY_VALOR_I": paymentCode = "MEDAL_MILITARY_VALOR_I": amountValue = 20
        Case "MILITARY_VALOR_II": paymentCode = "MEDAL_MILITARY_VALOR_II": amountValue = 10
        Case Else: Exit Sub
    End Select
    awardDate = TextValue(ruleData, "medal_award_date")
    If Not IsDate(awardDate) Then
        results.Add NewAllowance("Medal", paymentCode, "PERCENT", amountValue, "SPECIAL_ACHIEVEMENTS_P2", ALLOWANCE_STATUS_REQUIRES_DECISION, "Medal payment requires the award-order date and reference.", ruleData)
        SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 2 Правил"
        Exit Sub
    End If
    results.Add NewAllowance("Medal", paymentCode, "PERCENT", amountValue, "SPECIAL_ACHIEVEMENTS_P2", ALLOWANCE_STATUS_ACTIVE, "Medal payment is payable for one year from the award-order date.", ruleData)
    Set allowance = results(results.Count)
    allowance("start_date") = CDate(awardDate)
    allowance("end_date") = DateAdd("yyyy", 1, CDate(awardDate)) - 1
    allowance("document_reference") = TextValue(ruleData, "medal_award_document_reference")
    allowance("factual_basis") = "Medal " & medalCode & "; award date " & Format$(CDate(awardDate), "dd.mm.yyyy")
    SetAllowanceAct allowance, mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 2 Правил"
End Sub

Private Sub AddVusResult(ByVal results As Collection, ByVal stateData As Object, ByVal ruleData As Object)
    Dim vusCode As String
    vusCode = UCase$(Replace$(TextValue(stateData, "vus"), "-", ""))
    If vusCode = "310100" Or vusCode = "310101" Then
        results.Add NewAllowance("VUS", "VUS_310100_310101", "PERCENT", 50, "SPECIAL_ACHIEVEMENTS_P2", ALLOWANCE_STATUS_ACTIVE, "VUS 310100 or 310101.", ruleData)
        SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 2 Правил"
    End If
End Sub

Private Sub AddDriverResult(ByVal results As Collection, ByVal stateData As Object, ByVal ruleData As Object)
    If IsTrueValue(TextValue(stateData, "driver_c_d_ce")) Then
        results.Add NewAllowance("Driver", "DRIVER_C_D_CE", "PERCENT", 30, "SPECIAL_ACHIEVEMENTS_P2", ALLOWANCE_STATUS_ACTIVE, "Driver position and C/D/CE entitlement require supporting documents.", ruleData)
        SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 2 Правил"
    End If
End Sub

Private Sub AddTariffResult(ByVal results As Collection, ByVal stateData As Object, ByVal ruleData As Object)
    Dim tariffRank As Long
    If IsNumeric(TextValue(stateData, "tariff_rank")) Then
        tariffRank = CLng(TextValue(stateData, "tariff_rank"))
        If tariffRank >= 1 And tariffRank <= 4 Then
            results.Add NewAllowance("Tariff rank", "TARIFF_1_4", "PERCENT", 50, "POINT_3_SEPARATE", ALLOWANCE_STATUS_ACTIVE, "Point 3 allowance; it is outside the point-2 cap.", ruleData)
            SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 3 Правил"
        End If
    End If
End Sub

Private Sub AddContract430Result(ByVal results As Collection, ByVal stateData As Object, ByVal ruleData As Object)
    If IsTrueValue(TextValue(stateData, "contract_430_eligible")) Then
        results.Add NewAllowance("Contract allowance", "MOBILIZATION_OR_SVO_CONTRACT", "PERCENT", 60, "POINT_3_4_SEPARATE", ALLOWANCE_STATUS_ACTIVE, "Point 3.4 allowance; eligibility must be evidenced by the applicable contract or mobilization documents.", ruleData)
        SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_MO_430, "пункт 3.4 Правил"
    End If
End Sub

Private Sub ApplyPoint2Cap(ByVal results As Collection, ByRef point2Total As Double)
    Dim allowance As Object
    Dim activeCount As Long
    Dim capPercent As Double

    point2Total = 0
    For Each allowance In results
        If TextValue(allowance, "cap_group") = "SPECIAL_ACHIEVEMENTS_P2" And TextValue(allowance, "status") = ALLOWANCE_STATUS_ACTIVE Then
            point2Total = point2Total + NumericValue(allowance, "amount_value")
            activeCount = activeCount + 1
        End If
    Next allowance

    capPercent = GetConfiguredCapPercent("SPECIAL_ACHIEVEMENTS_P2", 100)
    If point2Total <= capPercent Then Exit Sub
    For Each allowance In results
        If TextValue(allowance, "cap_group") = "SPECIAL_ACHIEVEMENTS_P2" And TextValue(allowance, "status") = ALLOWANCE_STATUS_ACTIVE Then
            allowance("applied_amount") = allowance("original_amount")
            allowance("explanation") = "Point-2 grounds total " & Format$(point2Total, "0.##") & "%. The group is payable in total no more than " & Format$(capPercent, "0.##") & "%; retain every original ground and percentage in the order."
        End If
    Next allowance
End Sub

Private Function GetConfiguredCapPercent(ByVal capGroup As String, ByVal fallbackPercent As Double) As Double
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long

    GetConfiguredCapPercent = fallbackPercent
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PaymentCaps")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(rowNum, 1).Value))) = UCase$(capGroup) Then
            If UCase$(Trim$(CStr(ws.Cells(rowNum, 9).Value))) <> "INACTIVE" And IsNumeric(ws.Cells(rowNum, 2).Value) Then GetConfiguredCapPercent = CDbl(ws.Cells(rowNum, 2).Value)
            Exit Function
        End If
    Next rowNum
End Function

Private Sub AddMobilizedFixedPayment(ByVal results As Collection, ByVal stateData As Object, ByVal ruleData As Object, ByVal serviceCategory As String)
    If serviceCategory <> "MOBILIZED" Then Exit Sub

    results.Add NewAllowance("Mobilized social payment", "MOBILIZED_FIXED_158000", "FIXED_AMOUNT", 158000, "SEPARATE_LEGAL_ACT", ALLOWANCE_STATUS_ACTIVE, "Monthly social payment for mobilized citizens under Presidential Decree No. 788; procedure under Ministry of Defence Order No. 780.", ruleData)
    SetAllowanceAct results(results.Count), mdlPersonnelEvents.LEGAL_ACT_UP_788, "пункт 1; порядок — приказ МО № 780"
    results(results.Count)("start_date") = ResolveMobilizedFixedStartDate(ruleData)
End Sub

Private Sub SetAllowanceAct(ByVal allowance As Object, ByVal actID As String, ByVal actPoint As String)
    allowance("act_id") = actID
    allowance("act_point") = actPoint
End Sub

Private Function ResolveMobilizedFixedStartDate(ByVal ruleData As Object) As Variant
    Dim candidateDate As Date
    candidateDate = DateSerial(2022, 9, 21)
    If IsDate(TextValue(ruleData, "default_start_date")) Then
        If CDate(TextValue(ruleData, "default_start_date")) > candidateDate Then candidateDate = CDate(TextValue(ruleData, "default_start_date"))
    End If
    ResolveMobilizedFixedStartDate = candidateDate
End Function

Private Function NewAllowance(ByVal paymentType As String, ByVal paymentCode As String, ByVal amountKind As String, ByVal amountValue As Double, ByVal capGroup As String, ByVal statusValue As String, ByVal explanation As String, ByVal ruleData As Object) As Object
    Dim allowance As Object
    Set allowance = CreateObject("Scripting.Dictionary")
    allowance("payment_type") = paymentType
    allowance("payment_code") = paymentCode
    allowance("amount_kind") = amountKind
    allowance("amount_value") = CStr(amountValue)
    allowance("original_amount") = CStr(amountValue)
    allowance("applied_amount") = CStr(amountValue)
    allowance("cap_group") = capGroup
    allowance("status") = statusValue
    allowance("explanation") = explanation
    allowance("act_id") = TextValue(ruleData, LCase$(paymentCode) & "_act_id")
    allowance("start_date") = TextValue(ruleData, LCase$(paymentCode) & "_start_date")
    allowance("end_date") = TextValue(ruleData, LCase$(paymentCode) & "_end_date")
    allowance("document_reference") = TextValue(ruleData, LCase$(paymentCode) & "_document_reference")
    Set NewAllowance = allowance
End Function

Private Function TextValue(ByVal source As Object, ByVal key As String) As String
    If source Is Nothing Then Exit Function
    If source.Exists(key) Then TextValue = Trim$(CStr(source(key)))
End Function

Private Function NumericValue(ByVal source As Object, ByVal key As String) As Double
    If IsNumeric(TextValue(source, key)) Then NumericValue = CDbl(TextValue(source, key))
End Function

Private Function IsTrueValue(ByVal valueText As String) As Boolean
    valueText = UCase$(Trim$(valueText))
    IsTrueValue = (valueText = "YES" Or valueText = "TRUE" Or valueText = "1")
End Function
