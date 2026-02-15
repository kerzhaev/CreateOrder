VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "О программе"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Флаг для предотвращения "зацикливания" события Change
Private bIgnoreChange As Boolean

Private Sub UserForm_Initialize()
    ' Инициализация данных о программе
    lblProductName.Caption = PRODUCT_NAME
    lblVersion.Caption = "Версия: " & PRODUCT_VERSION
    lblAuthor.Caption = "Автор: " & PRODUCT_AUTHOR
    lblEmail.Caption = "E-mail: " & PRODUCT_EMAIL
    lblPhone.Caption = "Телефон: " & PRODUCT_PHONE
    lblCompany.Caption = "Организация: " & PRODUCT_COMPANY
    
    ' Настройка интерфейса в зависимости от статуса
    UpdateLicenseStatusUI
End Sub

Private Sub btnActivate_Click()
    Dim key As String
    key = Trim(txtActivationCode.Text)
    
    If key = "" Then
        MsgBox "Введите ключ активации.", vbExclamation
        Exit Sub
    End If
    
    ' Вызов функции проверки из модуля modActivation
    If modActivation.ActivateProduct(key) Then
        ' Если успешно - обновляем интерфейс
        UpdateLicenseStatusUI
        txtActivationCode.Text = ""
    End If
End Sub

Private Sub UpdateLicenseStatusUI()
    Dim status As Integer
    status = modActivation.GetLicenseStatus()
    
    If status = 0 Then
        ' Активна
        lblActivationStatus.Caption = "СТАТУС: АКТИВИРОВАНО (до " & modActivation.GetLicenseExpiryDateStr() & ")"
        lblActivationStatus.ForeColor = RGB(0, 150, 0) ' Зеленый
        
        ' Скрываем поля ввода, они больше не нужны
        txtActivationCode.Enabled = False
        btnActivate.Enabled = False
        lblActivationHint.Caption = "Продукт активирован. Спасибо!"
    Else
        ' Истекла или нет
        lblActivationStatus.Caption = "СТАТУС: ТРЕБУЕТСЯ АКТИВАЦИЯ"
        lblActivationStatus.ForeColor = RGB(200, 0, 0) ' Красный
        
        txtActivationCode.Enabled = True
        btnActivate.Enabled = True
        lblActivationHint.Caption = modActivation.ACTIVATION_HINT
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' Сброс лицензии (для отладки, можно назначить на секретную кнопку или убрать)
Private Sub lblVersion_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim response As VbMsgBoxResult
    response = MsgBox("Сбросить лицензию (для тестов)?", vbYesNo + vbQuestion)
    If response = vbYes Then
        On Error Resume Next
        ThisWorkbook.Names("LicData").Delete
        ThisWorkbook.Names("LicSign").Delete
        MsgBox "Лицензия сброшена."
        UpdateLicenseStatusUI
    End If
End Sub

' === АВТОМАТИЧЕСКОЕ ФОРМАТИРОВАНИЕ КЛЮЧА (XXXX-XXXX-XXXX-XXXX) ===



Private Sub txtActivationCode_Change()
    ' Если изменение вызвано нашим кодом, ничего не делаем
    If bIgnoreChange Then Exit Sub
    
    Dim rawText As String
    Dim cleanText As String
    Dim formattedText As String
    Dim i As Integer
    
    ' Включаем блокировку событий
    bIgnoreChange = True
    
    ' 1. Получаем текущий текст и сразу переводим в верхний регистр
    rawText = UCase(Me.txtActivationCode.Text)
    
    ' 2. Очищаем от всего лишнего (тире, пробелы), оставляем только буквы/цифры
    cleanText = Replace(Replace(rawText, "-", ""), " ", "")
    
    ' 3. Ограничиваем длину ввода (максимум 16 символов самого ключа)
    If Len(cleanText) > 16 Then cleanText = Left(cleanText, 16)
    
    ' 4. Собираем строку заново, вставляя тире
    formattedText = ""
    For i = 1 To Len(cleanText)
        formattedText = formattedText & Mid(cleanText, i, 1)
        
        ' Вставляем тире после каждого 4-го символа (4, 8, 12), но не в самом конце
        If (i Mod 4 = 0) And (i < 16) Then
            formattedText = formattedText & "-"
        End If
    Next i
    
    ' 5. Записываем отформатированный текст обратно
    Me.txtActivationCode.Text = formattedText
    
    ' 6. Ставим курсор в конец строки (чтобы при вставке в середину курсор не прыгал, можно усложнить, но для ввода ключа обычно вводят последовательно)
    Me.txtActivationCode.SelStart = Len(formattedText)
    
    ' Снимаем блокировку
    bIgnoreChange = False
End Sub

' Дополнительно: Запрет на ввод русских букв и спецсимволов (только латиница и цифры)
Private Sub txtActivationCode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim char As String
    
    ' Используем ChrW вместо Chr, чтобы не было ошибки на русских буквах (Unicode)
    char = UCase(ChrW(KeyAscii))
    
    ' Разрешаем: 0-9, A-Z, Backspace (8)
    ' Запрещаем все остальное (включая русские буквы)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", char) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0 ' Отменяем ввод символа
    Else
        ' Принудительно переводим в верхний регистр при вводе
        ' (Если введена маленькая латинская буква a-z)
        If KeyAscii >= 97 And KeyAscii <= 122 Then
            KeyAscii = KeyAscii - 32
        End If
    End If
End Sub
