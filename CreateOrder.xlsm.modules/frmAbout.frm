VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "UserForm1"
   ClientHeight    =   9045.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10545
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' frmAbout - Форма информации и активации макроса
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================================

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As String) As LongPtr
#Else
    Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "User32" () As Long
    Private Declare Function EmptyClipboard Lib "User32" () As Long
    Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
#End If

Const GHND = &H42
Const CF_TEXT = 1



Private Sub UserForm_Initialize()

    Call mdlHelper.EnsureStaffColumnsInitialized

    lblProductName.Caption = PRODUCT_NAME
    lblVersion.Caption = "Версия: " & PRODUCT_VERSION
    lblAuthor.Caption = "Автор: " & PRODUCT_AUTHOR
    lblEmail.Caption = "E-mail: " & PRODUCT_EMAIL
    lblPhone.Caption = "Телефон: " & PRODUCT_PHONE
    lblCompany.Caption = "Организация: " & PRODUCT_COMPANY
    lblActivationHint.Caption = ACTIVATION_HINT

    lblActivationStatus.Caption = GetTrialStatusText
    If GetProductStatus() = 0 Then
        lblActivationStatus.ForeColor = vbGreen
        txtActivationCode.Enabled = False
        btnActivate.Enabled = False
    ElseIf GetProductStatus() = 1 Then
        lblActivationStatus.ForeColor = &H80FF   ' синий — триал
        txtActivationCode.Enabled = True
        btnActivate.Enabled = True
    Else
        lblActivationStatus.ForeColor = vbRed
        txtActivationCode.Enabled = True
        btnActivate.Enabled = True
    End If
End Sub

Private Sub btnActivate_Click()
    If ValidateActivationKey(txtActivationCode.text) Then
        SaveActivationStatus True
        lblActivationStatus.Caption = GetTrialStatusText
        lblActivationStatus.ForeColor = vbGreen
        txtActivationCode.Enabled = False
        btnActivate.Enabled = False
        MsgBox "Спасибо! Продукт активирован.", vbInformation
    Else
        MsgBox "Код активации неверный.", vbExclamation
    End If
End Sub


Private Sub btnCopyContact_Click()
    ' Копирует информацию для связи в буфер обмена
    Dim contactText As String
    contactText = "Автор: " & PRODUCT_AUTHOR & vbCrLf & _
                  "Email: " & PRODUCT_EMAIL & vbCrLf & _
                  "Телефон: " & PRODUCT_PHONE & vbCrLf & _
                  "Компания: " & PRODUCT_COMPANY
    PutToClipboard contactText
    MsgBox "Контактная информация скопирована в буфер обмена.", vbInformation
End Sub

Private Sub btnShowHelp_Click()
    ' Открытие справки (можно заменить вызов на собственную инструкцию)
    MsgBox "Здесь будет ваша инструкция/FAQ.", vbInformation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ===== Служебная процедура для копирования текста в буфер обмена =====
Private Sub PutToClipboard(ByVal text As String)
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    Dim hWnd As LongPtr
    hWnd = 0
    If OpenClipboard(hWnd) Then
        EmptyClipboard
        hGlobalMemory = GlobalAlloc(GHND, Len(text) + 1)
        lpGlobalMemory = GlobalLock(hGlobalMemory)
        lstrcpy lpGlobalMemory, text
        GlobalUnlock hGlobalMemory
        SetClipboardData CF_TEXT, hGlobalMemory
        CloseClipboard
    End If
End Sub


