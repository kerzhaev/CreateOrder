Attribute VB_Name = "modActivation"
'==============================================================
' Licensing with 365-day trial period and activation
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
'==============================================================
Option Explicit

Public Const PRODUCT_NAME As String = "Формирователь приказов"
Public Const PRODUCT_VERSION As String = "1.4 от 01.12.25"
Public Const PRODUCT_AUTHOR As String = "Кержаев Евгений Алексеевич"
Public Const PRODUCT_EMAIL As String = "nachfin@vk.com"
Public Const PRODUCT_PHONE As String = "+7(989)906-88-91"
Public Const PRODUCT_COMPANY As String = "95 ФЭС"
Public Const ACTIVATION_HINT As String = "Введите полученный код активации..."
Public Const MASTER_KEY As String = "TESTKEY-1234"
Public Const TRIAL_PERIOD_DAYS As Long = 365 ' Trial length (1 year)
Public Const TRIAL_NAME As String = "TrialStartDate" ' Hidden parameter name

'=== Check activation code ===
Public Function ValidateActivationKey(ByVal userKey As String) As Boolean
    If Trim(userKey) = MASTER_KEY Or Trim(userKey) = "DEVKEY-5678" Then
        ValidateActivationKey = True
    Else
        ValidateActivationKey = False
    End If
End Function

'=== Save activation status ===
Public Sub SaveActivationStatus(ByVal activated As Boolean)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="ActivatedProduct", RefersTo:="=" & activated
End Sub

'=== Load activation status ===
Public Function LoadActivationStatus() As Boolean
    On Error Resume Next
    LoadActivationStatus = CBool(ThisWorkbook.Names("ActivatedProduct").RefersTo)
End Function

'=== Reset status for debugging ===
Public Sub ResetActivationStatus()
    On Error Resume Next
    ThisWorkbook.Names("ActivatedProduct").Delete
End Sub

'=== First launch date (trial) ===
Public Function GetTrialStartDate() As Date
    On Error Resume Next
    Dim trialDate As Date
    trialDate = 0
    If Not NameExists(TRIAL_NAME) Then
        trialDate = Date
        ThisWorkbook.Names.Add Name:=TRIAL_NAME, RefersTo:=trialDate
    Else
        trialDate = ThisWorkbook.Names(TRIAL_NAME).RefersTo
    End If
    GetTrialStartDate = trialDate
End Function

'=== Check trial/license ===
' 0 - activated; 1 - trial; 2 - trial expired
Public Function GetProductStatus() As Integer
    If LoadActivationStatus Then
        GetProductStatus = 0
    Else
        Dim trialDate As Date
        trialDate = GetTrialStartDate()
        If DateDiff("d", trialDate, Date) < TRIAL_PERIOD_DAYS Then
            GetProductStatus = 1
        Else
            GetProductStatus = 2
        End If
    End If
End Function

'=== Get status string ===
Public Function GetTrialStatusText() As String
    Dim stat As Integer, trialDate As Date, daysLeft As Long
    stat = GetProductStatus()
    If stat = 0 Then
        GetTrialStatusText = "Статус: АКТИВИРОВАНО"
    ElseIf stat = 1 Then
        trialDate = GetTrialStartDate()
        daysLeft = TRIAL_PERIOD_DAYS - DateDiff("d", trialDate, Date)
        GetTrialStatusText = "Статус: Триал (" & daysLeft & " дн. осталось)"
    Else
        GetTrialStatusText = "Статус: Триал истёк — требуется активация"
    End If
End Function

'=== Check if name exists ===
Public Function NameExists(nm As String) As Boolean
    On Error Resume Next
    NameExists = Not ThisWorkbook.Names(nm) Is Nothing
End Function

'=== Reset trial (test/dev only!) ===
Public Sub ResetTrialDate()
    On Error Resume Next
    ThisWorkbook.Names(TRIAL_NAME).Delete
End Sub
