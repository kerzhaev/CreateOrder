VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "О программе"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===============================================================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Код формы "О программе" (Версия с точными именами контролов)
' ===============================================================================
Option Explicit

Private bIgnoreChange As Boolean

Private Sub UserForm_Initialize()
    ' Инициализация данных о программе
    lblProductName.Caption = modActivation.PRODUCT_NAME
    lblVersion.Caption = "Версия: " & modActivation.PRODUCT_VERSION
    lblAuthor.Caption = "Автор: " & modActivation.PRODUCT_AUTHOR
    lblEmail.Caption = "E-mail: " & modActivation.PRODUCT_EMAIL
    lblPhone.Caption = "Телефон: " & modActivation.PRODUCT_PHONE
    lblCompany.Caption = "Организация: " & modActivation.PRODUCT_COMPANY
    
    ' Обновляем интерфейс
    UpdateLicenseStatusUI
End Sub

' ПАСХАЛКА: Двойной клик по автору открывает панель администратора
Private Sub lblAuthor_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call modActivation.AdminSetGlobalDate
    UpdateLicenseStatusUI
End Sub

' Кнопка активации
Private Sub btnActivate_Click()
    Dim key As String
    key = Trim(txtActivationCode.Text)
    
    If key = "" Then
        MsgBox "Введите ключ активации.", vbExclamation
        Exit Sub
    End If
    
    ' Вызов функции проверки из модуля modActivation
    If modActivation.ActivatePersonal(key) Then
        MsgBox "Программа успешно активирована для данного ПК!", vbInformation
        UpdateLicenseStatusUI
        txtActivationCode.Text = ""
    Else
        MsgBox "Ключ не подходит. Проверьте правильность ввода или ID оборудования.", vbCritical
    End If
End Sub

' =============================================
' Обновление интерфейса в зависимости от статуса
' =============================================
Private Sub UpdateLicenseStatusUI()
    Dim status As Integer
    status = modActivation.GetLicenseStatus()
    
    Select Case status
        Case 0 ' ПЕРСОНАЛЬНАЯ ЛИЦЕНЗИЯ (По ключу HWID)
            lblActivationStatus.Caption = "СТАТУС: ПЕРСОНАЛЬНАЯ ЛИЦЕНЗИЯ"
            lblActivationStatus.ForeColor = RGB(0, 150, 0) ' Зеленый
            lblPremiumMessage.Caption = "Лицензия надежно привязана к этому ПК." & vbCrLf & _
                                        "Действует до: " & modActivation.GetLicenseExpiryDateStr()
            lblPremiumMessage.ForeColor = RGB(0, 100, 0)
            
            ' Скрываем поля, так как активация уже выполнена
            txtActivationCode.Visible = False
            btnActivate.Visible = False
            lblActivationHint.Visible = False
            
        Case 3 ' ГЛОБАЛЬНАЯ ЛИЦЕНЗИЯ (Через пасхалку админа)
            lblActivationStatus.Caption = "СТАТУС: КОРПОРАТИВНАЯ ВЕРСИЯ"
            lblActivationStatus.ForeColor = RGB(0, 150, 0) ' Зеленый
            lblPremiumMessage.Caption = "Файл предварительно активирован администратором." & vbCrLf & _
                                        "Действует до: " & modActivation.GetLicenseExpiryDateStr()
            lblPremiumMessage.ForeColor = RGB(0, 100, 0)
            
            ' Скрываем поля
            txtActivationCode.Visible = False
            btnActivate.Visible = False
            lblActivationHint.Visible = False
            
        Case 4 ' ПУБЛИЧНАЯ ВЕРСИЯ (Бесплатный период)
            lblActivationStatus.Caption = "СТАТУС: ОЗНАКОМИТЕЛЬНЫЙ ПЕРИОД"
            lblActivationStatus.ForeColor = RGB(200, 100, 0) ' Оранжевый
            lblPremiumMessage.Caption = "Ознакомительная версия активна до: " & modActivation.GetLicenseExpiryDateStr() & vbCrLf & _
                                        "ВАШ КОД ОБОРУДОВАНИЯ: " & modActivation.GetHardwareID() & vbCrLf & _
                                        "Если у вас уже есть персональный ключ активируйте его ниже, либо обратитесь для его получения по контактам, указанным выше:"
            lblPremiumMessage.ForeColor = RGB(100, 50, 0)
            
            ' ПОКАЗЫВАЕМ ПОЛЯ для заблаговременной активации!
            txtActivationCode.Visible = True
            btnActivate.Visible = True
            lblActivationHint.Visible = True
            lblActivationHint.Caption = modActivation.ACTIVATION_HINT
            
        Case 2 ' ПЕРЕВОД ЧАСОВ НАЗАД
            lblActivationStatus.Caption = "ОШИБКА: ОБНАРУЖЕН ПЕРЕВОД ЧАСОВ"
            lblActivationStatus.ForeColor = RGB(200, 0, 0) ' Красный
            lblPremiumMessage.Caption = "Защитная блокировка." & vbCrLf & _
                                        "КОД ОБОРУДОВАНИЯ: " & modActivation.GetHardwareID() & vbCrLf & _
                                        "Исправьте дату или введите ключ."
            lblPremiumMessage.ForeColor = RGB(200, 0, 0)
            
            txtActivationCode.Visible = True
            btnActivate.Visible = True
            lblActivationHint.Visible = True
            lblActivationHint.Caption = modActivation.ACTIVATION_HINT
            
        Case Else ' ИСТЕКЛО ИЛИ НЕТ ЛИЦЕНЗИИ (1)
            lblActivationStatus.Caption = "СТАТУС: ОГРАНИЧЕННАЯ ВЕРСИЯ"
            lblActivationStatus.ForeColor = RGB(200, 0, 0) ' Красный
            lblPremiumMessage.Caption = "ОЗНАКОМИТЕЛЬНЫЙ период завершен." & vbCrLf & _
                                        "ВАШ КОД ОБОРУДОВАНИЯ: " & modActivation.GetHardwareID() & vbCrLf & _
                                        "Для продолжения работы запросите по контактам указанным выше, персональный ключ."
            lblPremiumMessage.ForeColor = RGB(50, 50, 50)
            
            txtActivationCode.Visible = True
            btnActivate.Visible = True
            lblActivationHint.Visible = True
            lblActivationHint.Caption = modActivation.ACTIVATION_HINT
            
    End Select
End Sub

' =============================================
' АВТОМАТИЧЕСКОЕ ФОРМАТИРОВАНИЕ КЛЮЧА
' =============================================
Private Sub txtActivationCode_Change()
    If bIgnoreChange Then Exit Sub
    Dim rawText As String, cleanText As String, formattedText As String
    Dim i As Integer
    
    bIgnoreChange = True
    rawText = UCase(Me.txtActivationCode.Text)
    cleanText = Replace(Replace(rawText, "-", ""), " ", "")
    
    If Len(cleanText) > 16 Then cleanText = Left(cleanText, 16)
    
    formattedText = ""
    For i = 1 To Len(cleanText)
        formattedText = formattedText & Mid(cleanText, i, 1)
        If (i Mod 4 = 0) And (i < 16) Then formattedText = formattedText & "-"
    Next i
    
    Me.txtActivationCode.Text = formattedText
    Me.txtActivationCode.SelStart = Len(formattedText)
    bIgnoreChange = False
End Sub

' Запрет кириллицы
Private Sub txtActivationCode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim char As String
    char = UCase(ChrW(KeyAscii))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", char) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    Else
        If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

