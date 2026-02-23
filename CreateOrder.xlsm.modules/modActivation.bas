Attribute VB_Name = "modActivation"
' ===============================================================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Гибридная система лицензирования (Trial + Global Admin + Personal HWID)
' @version 2.1.0
' ===============================================================================
Option Explicit

' --- ИНФОРМАЦИЯ О ПРОДУКТЕ (ДЛЯ ФОРМЫ FRMABOUT) ---
Public Const PRODUCT_NAME As String = "Формирователь приказов"
Public Const PRODUCT_VERSION As String = "2.1.0"
Public Const PRODUCT_AUTHOR As String = "Кержаев Евгений Алексеевич"
Public Const PRODUCT_EMAIL As String = "nachfin@vk.com"
Public Const PRODUCT_PHONE As String = "+7(989)906-88-91"
Public Const PRODUCT_COMPANY As String = "Отделение по разработке программного обеспечения 95 ФЭС МО РФ"
Public Const ACTIVATION_HINT As String = "Введите ключ (формат: XXXX-XXXX-XXXX-XXXX)"

' --- ГЛОБАЛЬНЫЕ НАСТРОЙКИ ЗАЩИТЫ ---
Private Const ADMIN_PASSWORD As String = "95FES_Admin"    ' Пароль для "Пасхалки"
Private Const SALT_KEY As String = "SECURE_TOKEN_2026"    ' Соль для хэширования
Private Const MAGIC_SEED As Long = 1985                   ' Смещение для шифрования дат

' --- ИМЕНА СКРЫТЫХ ДИАПАЗОНОВ (ПАМЯТЬ КНИГИ EXCEL) ---
Private Const NAME_PERSONAL_DATA As String = "LicData"    ' Зашифрованная дата персонального ключа
Private Const NAME_PERSONAL_SIGN As String = "LicSign"    ' Цифровая подпись (привязана к HWID)
Private Const NAME_GLOBAL_DATA As String = "GlobalLimit"  ' Дата, установленная администратором
Private Const NAME_GLOBAL_SIGN As String = "GlobalSign"   ' Подпись даты администратора (без HWID)
Private Const NAME_LAST_RUN As String = "LicLast"         ' Дата последнего запуска (от перевода часов)

' ===============================================================================
' 1. ИДЕНТИФИКАЦИЯ ОБОРУДОВАНИЯ (HWID)
' ===============================================================================
Public Function GetHardwareID() As String
    On Error Resume Next
    Dim fso As Object, d As Object, hexStr As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive("C:\")
    
    If Err.number = 0 Then
        hexStr = Hex(d.SerialNumber)
        ' Дополняем нулями до 8 символов
        GetHardwareID = Right("00000000" & hexStr, 8)
    End If
    
    If GetHardwareID = "00000000" Or GetHardwareID = "" Then GetHardwareID = "NODEFAULT"
    On Error GoTo 0
End Function

' ===============================================================================
' 2. ГЕНЕРАЦИЯ ПЕРСОНАЛЬНОГО КЛЮЧА (ДЛЯ АДМИНА)
' Вводить в Immediate Window: ?modActivation.GenerateLicenseKey("31.12.2026", "A1B2C3D4")
' ===============================================================================
Public Function GenerateLicenseKey(expiryDate As Date, targetHWID As String) As String
    Dim p1_Salt As String, p2_Date As String, p3_Noise As String, p4_Check As String
    
    Randomize
    p1_Salt = Right("0000" & Hex(Int((65535 * Rnd) + 1)), 4)
    p2_Date = Right("0000" & Hex(CLng(expiryDate) - MAGIC_SEED Xor CLng("&H" & p1_Salt)), 4)
    p3_Noise = Right("0000" & Hex(Int((65535 * Rnd) + 1)), 4)
    p4_Check = CalculateHash(p1_Salt & p2_Date & p3_Noise, UCase(Trim(targetHWID)))
    
    GenerateLicenseKey = p1_Salt & "-" & p2_Date & "-" & p3_Noise & "-" & p4_Check
End Function

' ===============================================================================
' 3. АКТИВАЦИЯ ПЕРСОНАЛЬНОГО КЛЮЧА
' ===============================================================================
Public Function ActivatePersonal(key As String) As Boolean
    Dim p() As String, dVal As Long, d As Date, localHWID As String
    
    ' Очищаем и разбиваем ключ
    key = UCase(Replace(Trim(key), " ", ""))
    p = Split(key, "-")
    
    If UBound(p) <> 3 Then
        ActivatePersonal = False
        Exit Function
    End If
    
    localHWID = GetHardwareID()
    
    ' Проверяем подпись ключа с учетом локального HWID
    If p(3) <> CalculateHash(p(0) & p(1) & p(2), localHWID) Then
        ActivatePersonal = False
        Exit Function
    End If
    
    ' Расшифровываем дату
    On Error Resume Next
    dVal = CLng("&H" & p(1)) Xor CLng("&H" & p(0))
    d = CDate(dVal + MAGIC_SEED)
    
    If Err.number <> 0 Then
        ActivatePersonal = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Записываем в скрытую память с привязкой к HWID
    WriteHidden NAME_PERSONAL_DATA, CStr(CLng(d))
    WriteHidden NAME_PERSONAL_SIGN, CalculateHash(CStr(CLng(d)), localHWID)
    
    ActivatePersonal = True
End Function

' ===============================================================================
' 4. ГЛАВНАЯ ЛОГИКА ПРОВЕРКИ ЛИЦЕНЗИИ
' Возвращает:
' 0 - Активно (Персональный ключ)
' 3 - Активно (Глобальная дата админа)
' 4 - Активно (Бесплатный период / Публичная)
' 1 - Истекло / Нет лицензии
' 2 - Взлом времени
' ===============================================================================
Public Function GetLicenseStatus() As Integer
    Dim personalExp As Date, globalExp As Date, publicExp As Date
    Dim lastRun As Date, lastRunStr As String
    
    On Error Resume Next
    
    ' ШАГ 1: Защита от перевода часов назад
    lastRunStr = ReadHidden(NAME_LAST_RUN)
    If lastRunStr <> "" Then
        lastRun = CDate(val(lastRunStr))
        If Date < lastRun And lastRun > 0 Then
            GetLicenseStatus = 2
            Exit Function
        End If
    End If
    
    ' ШАГ 2: Сбор всех доступных дат
    publicExp = DateSerial(2026, 6, 1) ' 1 Июня 2026 (Публичная триал-версия)
    globalExp = GetDateFromHidden(NAME_GLOBAL_DATA, NAME_GLOBAL_SIGN, False)
    personalExp = GetDateFromHidden(NAME_PERSONAL_DATA, NAME_PERSONAL_SIGN, True)
    
    ' Обновляем метку времени текущего запуска
    UpdateLastRun
    
    ' ШАГ 3: Иерархия (Персональная > Глобальная > Публичная)
    
    ' 1. Если есть персональный ключ (высший приоритет)
    If personalExp > 0 Then
        If Date > personalExp Then GetLicenseStatus = 1 Else GetLicenseStatus = 0
        Exit Function
    End If
    
    ' 2. Если есть дата от админа (пасхалка). Она полностью ПЕРЕБИВАЕТ публичную!
    If globalExp > 0 Then
        If Date > globalExp Then GetLicenseStatus = 1 Else GetLicenseStatus = 3
        Exit Function
    End If
    
    ' 3. Если ничего нет, работает публичная дата (Триал)
    If Date > publicExp Then
        GetLicenseStatus = 1
    Else
        GetLicenseStatus = 4 ' Код публичного триала
    End If
    
    On Error GoTo 0
End Function

' Получить строку с датой окончания лицензии для формы
Public Function GetLicenseExpiryDateStr() As String
    Dim personalExp As Date, globalExp As Date, publicExp As Date
    Dim finalExp As Date
    
    publicExp = DateSerial(2026, 6, 1) ' 1 Июня 2026
    globalExp = GetDateFromHidden(NAME_GLOBAL_DATA, NAME_GLOBAL_SIGN, False)
    personalExp = GetDateFromHidden(NAME_PERSONAL_DATA, NAME_PERSONAL_SIGN, True)
    
    ' Выбираем актуальную дату по приоритету
    finalExp = publicExp
    If globalExp > 0 Then finalExp = globalExp
    If personalExp > 0 Then finalExp = personalExp
    
    GetLicenseExpiryDateStr = Format(finalExp, "dd.mm.yyyy")
End Function

' Вызов перед выполнением макросов (от кнопок на ленте)
Public Function CheckLicenseAndPrompt() As Boolean
    Dim st As Integer
    st = GetLicenseStatus()
    
    ' Разрешаем работу для статусов 0, 3 и 4
    If st = 0 Or st = 3 Or st = 4 Then
        CheckLicenseAndPrompt = True
    Else
        MsgBox "Бесплатный период завершен. Для использования данной функции требуется действующая лицензия.", vbExclamation, "Блокировка"
        frmAbout.Show
        
        ' Повторная проверка после закрытия формы
        st = GetLicenseStatus()
        If st = 0 Or st = 3 Or st = 4 Then
            CheckLicenseAndPrompt = True
        Else
            CheckLicenseAndPrompt = False
        End If
    End If
End Function

' ===============================================================================
' 5. ПАСХАЛКА: ПАНЕЛЬ АДМИНИСТРАТОРА (Для выдачи файла без ключа)
' ===============================================================================
Public Sub AdminSetGlobalDate()
    Dim pwd As String, dStr As String, d As Date
    
    pwd = InputBox("Введите пароль администратора:", "Admin Panel")
    If pwd <> ADMIN_PASSWORD Then
        If pwd <> "" Then MsgBox "Неверный пароль!", vbCritical
        Exit Sub
    End If
    
    dStr = InputBox("Введите новую дату отключения для ЭТОГО ФАЙЛА (ДД.ММ.ГГГГ):" & vbCrLf & _
                    "Оставьте пустым для отмены.", "Установка Глобальной Лицензии", "01.06.2026")
                    
    If dStr = "" Then Exit Sub
    
    If IsDate(dStr) Then
        d = CDate(dStr)
        WriteHidden NAME_GLOBAL_DATA, CStr(CLng(d))
        WriteHidden NAME_GLOBAL_SIGN, CalculateHash(CStr(CLng(d)), "GLOBAL")
        MsgBox "Глобальная дата успешно установлена: " & Format(d, "dd.mm.yyyy"), vbInformation
    Else
        MsgBox "Некорректный формат даты.", vbExclamation
    End If
End Sub

' ===============================================================================
' 6. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (КРИПТОГРАФИЯ И ХРАНЕНИЕ)
' ===============================================================================
Private Function GetDateFromHidden(dataName As String, signName As String, useHWID As Boolean) As Date
    Dim dRaw As String, sRaw As String, expectedSign As String
    
    dRaw = ReadHidden(dataName)
    sRaw = ReadHidden(signName)
    
    If dRaw = "" Then
        GetDateFromHidden = 0
        Exit Function
    End If
    
    ' Формируем подпись для проверки
    If useHWID Then
        expectedSign = CalculateHash(dRaw, GetHardwareID())
    Else
        expectedSign = CalculateHash(dRaw, "GLOBAL")
    End If
    
    ' Сверяем подписи
    If sRaw = expectedSign Then
        On Error Resume Next
        GetDateFromHidden = CDate(CLng(dRaw))
        On Error GoTo 0
    Else
        GetDateFromHidden = 0
    End If
End Function

Private Function CalculateHash(rawStr As String, saltExtra As String) As String
    Dim i As Long, hash As Long
    Dim fullStr As String
    
    fullStr = rawStr & saltExtra & SALT_KEY
    hash = 0
    
    For i = 1 To Len(fullStr)
        hash = (hash * 31 + Asc(Mid(fullStr, i, 1))) Mod 65535
    Next i
    
    CalculateHash = Right("0000" & Hex(hash), 4)
End Function

Private Sub UpdateLastRun()
    WriteHidden NAME_LAST_RUN, CStr(CLng(Date))
End Sub

Private Sub WriteHidden(nName As String, nValue As String)
    On Error Resume Next
    ThisWorkbook.Names(nName).Delete
    ThisWorkbook.Names.Add Name:=nName, RefersTo:="=""" & nValue & """", Visible:=False
    On Error GoTo 0
End Sub

Private Function ReadHidden(nName As String) As String
    On Error Resume Next
    Dim v As String
    v = ThisWorkbook.Names(nName).RefersTo
    If Err.number = 0 Then
        v = Replace(v, "=", "")
        v = Replace(v, """", "")
        ReadHidden = v
    Else
        ReadHidden = ""
    End If
    On Error GoTo 0
End Function

