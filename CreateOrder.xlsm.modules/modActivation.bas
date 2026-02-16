Attribute VB_Name = "modActivation"
' ===============================================================================
' Модуль лицензирования (Оффлайн, с защитой от перевода часов)
' Формат ключа: AAAA-BBBB-CCCC-DDDD
' ===============================================================================
Option Explicit

' --- КОНСТАНТЫ ПРОДУКТА ---
Public Const PRODUCT_NAME As String = "Формирователь приказов"
Public Const PRODUCT_VERSION As String = "1.5.2"
Public Const PRODUCT_AUTHOR As String = "Кержаев Евгений Алексеевич"
Public Const PRODUCT_EMAIL As String = "nachfin@vk.com"
Public Const PRODUCT_PHONE As String = "+7(989)906-88-91"
Public Const PRODUCT_COMPANY As String = "95 ФЭС"
Public Const ACTIVATION_HINT As String = "Введите ключ (формат: XXXX-XXXX-XXXX-XXXX)"

' --- СЕКРЕТНЫЕ КОНСТАНТЫ (Внимание: измените их перед выпуском!) ---
' MAGIC_SEED - смещение даты, SALT_KEY - соль для хеширования
Private Const MAGIC_SEED As Long = 1985
Private Const SALT_KEY As String = "95FES_SECURE_2026"

' --- ИМЕНА СКРЫТЫХ ДИАПАЗОНОВ (Хранилище лицензии внутри файла) ---
Private Const NAME_LICENSE_DATA As String = "LicData"      ' Зашифрованная дата окончания
Private Const NAME_LICENSE_SIGN As String = "LicSign"      ' Подпись даты окончания
Private Const NAME_LAST_RUN As String = "LicLastRun"       ' Дата последнего запуска
Private Const NAME_LAST_RUN_SIGN As String = "LicLastRunS" ' Подпись последнего запуска

' ===============================================================================
' 1. ГЕНЕРАТОР КЛЮЧЕЙ (ДЛЯ АДМИНИСТРАТОРА)
' Вызывать из Immediate Window: ?GenerateLicenseKey("31.12.2026")
' ===============================================================================
Public Function GenerateLicenseKey(expiryDate As Date) As String
    Dim p1_Salt As String
    Dim p2_Date As String
    Dim p3_Noise As String
    Dim p4_Check As String
    Dim dateVal As Long
    
    Randomize
    
    ' 1. Соль (случайное число Hex)
    p1_Salt = Right("0000" & Hex(Int((65535 * Rnd) + 1)), 4)
    
    ' 2. Дата (XOR с солью + Смещение)
    dateVal = CLng(expiryDate) - MAGIC_SEED
    dateVal = dateVal Xor CLng("&H" & p1_Salt)
    p2_Date = Right("0000" & Hex(dateVal), 4)
    
    ' 3. Шум
    p3_Noise = Right("0000" & Hex(Int((65535 * Rnd) + 1)), 4)
    
    ' 4. Контрольная сумма (подпись первых трех частей)
    p4_Check = CalculateChecksum(p1_Salt & p2_Date & p3_Noise)
    
    GenerateLicenseKey = p1_Salt & "-" & p2_Date & "-" & p3_Noise & "-" & p4_Check
End Function

' ===============================================================================
' 2. ПРОВЕРКА КЛЮЧА И АКТИВАЦИЯ (ДЛЯ ПОЛЬЗОВАТЕЛЯ)
' ===============================================================================
Public Function ActivateProduct(key As String) As Boolean
    Dim parts() As String
    Dim p1 As String, p2 As String, p3 As String, p4 As String
    Dim expectedCheck As String
    Dim dateVal As Long
    Dim expiryDate As Date
    
    ' 1. Очистка и формат
    key = UCase(Replace(Trim(key), " ", ""))
    parts = Split(key, "-")
    
    If UBound(parts) <> 3 Then
        MsgBox "Неверный формат ключа.", vbExclamation
        ActivateProduct = False
        Exit Function
    End If
    
    p1 = parts(0): p2 = parts(1): p3 = parts(2): p4 = parts(3)
    
    ' 2. Проверка целостности
    expectedCheck = CalculateChecksum(p1 & p2 & p3)
    If p4 <> expectedCheck Then
        MsgBox "Ключ недействителен (ошибка контрольной суммы).", vbCritical
        ActivateProduct = False
        Exit Function
    End If
    
    ' 3. Расшифровка даты
    On Error Resume Next
    dateVal = CLng("&H" & p2)
    dateVal = dateVal Xor CLng("&H" & p1)
    dateVal = dateVal + MAGIC_SEED
    expiryDate = CDate(dateVal)
    
    If Err.number <> 0 Then
        MsgBox "Ключ поврежден.", vbCritical
        ActivateProduct = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' 4. Проверка срока
    If expiryDate < Date Then
        MsgBox "Срок действия этого ключа уже истёк (" & Format(expiryDate, "dd.mm.yyyy") & ").", vbExclamation
        ActivateProduct = False
        Exit Function
    End If
    
    ' 5. АКТИВАЦИЯ: Сохраняем лицензию и инициализируем таймер защиты
    SaveLicenseState expiryDate
    UpdateLastRunDate Date ' Устанавливаем текущую дату как "последний запуск"
    
    MsgBox "Программа успешно активирована!" & vbCrLf & _
           "Лицензия действует до: " & Format(expiryDate, "dd.mm.yyyy"), vbInformation
           
    ActivateProduct = True
End Function

' ===============================================================================
' 3. ПОЛУЧЕНИЕ СТАТУСА (С ПРОВЕРКОЙ ВРЕМЕНИ)
' Возвращает: 0 - Активна, 1 - Истекла/Нет, 2 - Блокировка (перевод часов)
' ===============================================================================
Public Function GetLicenseStatus() As Integer
    On Error Resume Next
    Dim licDateRaw As String, storedSign As String
    Dim lastRunRaw As String, lastRunSign As String
    Dim expDate As Date, lastRunDate As Date
    
    ' --- Шаг A: Читаем лицензию ---
    licDateRaw = ReadHiddenName(NAME_LICENSE_DATA)
    storedSign = ReadHiddenName(NAME_LICENSE_SIGN)
    
    ' Если лицензии нет или подпись не совпадает -> Статус 1
    If licDateRaw = "" Then GetLicenseStatus = 1: Exit Function
    If storedSign <> CalculateChecksum(licDateRaw) Then GetLicenseStatus = 1: Exit Function
    
    expDate = CDate(CLng(licDateRaw))
    
    ' --- Шаг B: Защита от перевода часов ---
    lastRunRaw = ReadHiddenName(NAME_LAST_RUN)
    lastRunSign = ReadHiddenName(NAME_LAST_RUN_SIGN)
    
    ' Если даты запуска нет (первый запуск после активации или сбой), считаем текущую дату
    If lastRunRaw = "" Or lastRunSign <> CalculateChecksum(lastRunRaw) Then
        lastRunDate = Date
        UpdateLastRunDate Date
    Else
        lastRunDate = CDate(CLng(lastRunRaw))
    End If
    
    ' ПРОВЕРКА 1: Перевод часов назад (Текущая дата < Последней сохраненной)
    If Date < lastRunDate Then
        MsgBox "ВНИМАНИЕ: Обнаружено изменение системного времени!" & vbCrLf & vbCrLf & _
               "Последний запуск: " & Format(lastRunDate, "dd.mm.yyyy") & vbCrLf & _
               "Текущая дата: " & Format(Date, "dd.mm.yyyy") & vbCrLf & vbCrLf & _
               "В целях безопасности лицензия временно заблокирована." & vbCrLf & _
               "Восстановите корректную дату на компьютере.", vbCritical, "Ошибка безопасности"
        GetLicenseStatus = 2 ' Блокировка
        Exit Function
    End If
    
    ' ПРОВЕРКА 2: Истечение срока (Текущая дата > Даты окончания)
    If Date > expDate Then
        GetLicenseStatus = 1 ' Истекла
        Exit Function
    End If
    
    ' --- Шаг C: Всё в порядке ---
    ' Если сегодня новый день (дата больше последней), обновляем метку последнего запуска
    If Date > lastRunDate Then
        UpdateLastRunDate Date
        ' Раскомментируйте следующую строку, если хотите принудительно сохранять файл для фиксации даты (надежнее, но навязчиво)
        ' ThisWorkbook.Save
    End If
    
    GetLicenseStatus = 0 ' Активна
End Function

' Получение строки с датой окончания (для формы About)
Public Function GetLicenseExpiryDateStr() As String
    On Error Resume Next
    Dim v As String
    v = ReadHiddenName(NAME_LICENSE_DATA)
    If IsNumeric(v) Then
        GetLicenseExpiryDateStr = Format(CDate(CLng(v)), "dd.mm.yyyy")
    Else
        GetLicenseExpiryDateStr = "Не активировано"
    End If
End Function

' ===============================================================================
' 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (ХЕШИРОВАНИЕ И ХРАНЕНИЕ)
' ===============================================================================

' Хеширование строки (для подписи)
Private Function CalculateChecksum(inputStr As String) As String
    Dim i As Long
    Dim hash As Long
    Dim fullStr As String
    
    ' Добавляем соль, чтобы нельзя было подделать подпись без знания кода
    fullStr = inputStr & SALT_KEY
    hash = 0
    
    ' Простой алгоритм хеширования
    For i = 1 To Len(fullStr)
        hash = (hash * 31 + Asc(Mid(fullStr, i, 1))) Mod 65535
    Next i
    
    CalculateChecksum = Right("0000" & Hex(hash), 4)
End Function

' Сохранение лицензии
Private Sub SaveLicenseState(expDate As Date)
    WriteHiddenName NAME_LICENSE_DATA, CStr(CLng(expDate))
    WriteHiddenName NAME_LICENSE_SIGN, CalculateChecksum(CStr(CLng(expDate)))
End Sub

' Обновление даты последнего запуска
Private Sub UpdateLastRunDate(runDate As Date)
    WriteHiddenName NAME_LAST_RUN, CStr(CLng(runDate))
    WriteHiddenName NAME_LAST_RUN_SIGN, CalculateChecksum(CStr(CLng(runDate)))
End Sub

' Запись в скрытое имя (Named Range)
Private Sub WriteHiddenName(nName As String, nValue As String)
    On Error Resume Next
    ' Удаляем старое имя, чтобы избежать конфликтов
    ThisWorkbook.Names(nName).Delete
    ' Создаем новое скрытое имя
    ThisWorkbook.Names.Add Name:=nName, RefersTo:="=""" & nValue & """", Visible:=False
End Sub

' Чтение из скрытого имени
Private Function ReadHiddenName(nName As String) As String
    On Error Resume Next
    Dim v As String
    v = ThisWorkbook.Names(nName).RefersTo
    ' Очистка мусора Excel (он возвращает формулу вида ="12345")
    v = Replace(v, "=", "")
    v = Replace(v, """", "")
    ReadHiddenName = v
End Function

