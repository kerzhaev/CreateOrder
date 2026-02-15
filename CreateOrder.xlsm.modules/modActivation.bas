Attribute VB_Name = "modActivation"
' ===============================================================================
' Модуль лицензирования (Оффлайн, с датой окончания)
' Формат ключа: AAAA-BBBB-CCCC-DDDD
' ===============================================================================
Option Explicit

' --- КОНСТАНТЫ ПРОДУКТА (ЭТО ТЕ, КОТОРЫХ НЕ ХВАТАЕТ) ---
Public Const PRODUCT_NAME As String = "Формирователь приказов"
Public Const PRODUCT_VERSION As String = "1.5.0"
Public Const PRODUCT_AUTHOR As String = "Кержаев Евгений Алексеевич"
Public Const PRODUCT_EMAIL As String = "nachfin@vk.com"
Public Const PRODUCT_PHONE As String = "+7(989)906-88-91"
Public Const PRODUCT_COMPANY As String = "95 ФЭС"
Public Const ACTIVATION_HINT As String = "Введите ключ (формат: XXXX-XXXX-XXXX-XXXX)"

' Секретные константы для шифрования (Измените их под себя!)
Private Const MAGIC_SEED As Long = 1985 ' Смещение даты
Private Const SALT_KEY As String = "95FES_SECURE_2026" ' Пароль для подписи

' Имена скрытых диапазонов для хранения статуса в файле Excel
Private Const NAME_LICENSE_DATA As String = "LicData"
Private Const NAME_LICENSE_SIGN As String = "LicSign"


' ===============================================================================
' ГЕНЕРАТОР КЛЮЧЕЙ (ДЛЯ АДМИНИСТРАТОРА)
' Вызывать из Immediate Window: ?GenerateLicenseKey("31.12.2026")
' ===============================================================================
Public Function GenerateLicenseKey(expiryDate As Date) As String
    Dim p1_Salt As String
    Dim p2_Date As String
    Dim p3_Noise As String
    Dim p4_Check As String
    Dim dateVal As Long
    
    Randomize
    
    ' 1. Группа 1: Случайная соль (Hex, 4 символа)
    p1_Salt = Right("0000" & Hex(Int((65535 * Rnd) + 1)), 4)
    
    ' 2. Группа 2: Дата (Смещение + XOR с солью для привязки)
    dateVal = CLng(expiryDate) - MAGIC_SEED
    ' Простое шифрование: XOR даты с первой частью ключа (преобразованной в число)
    dateVal = dateVal Xor CLng("&H" & p1_Salt)
    p2_Date = Right("0000" & Hex(dateVal), 4)
    
    ' 3. Группа 3: Шум (Просто случайные символы)
    p3_Noise = Right("0000" & Hex(Int((65535 * Rnd) + 1)), 4)
    
    ' 4. Группа 4: Контрольная сумма (Подпись первых трех частей)
    p4_Check = CalculateChecksum(p1_Salt & p2_Date & p3_Noise)
    
    GenerateLicenseKey = p1_Salt & "-" & p2_Date & "-" & p3_Noise & "-" & p4_Check
End Function

' ===============================================================================
' ПРОВЕРКА КЛЮЧА (ДЛЯ ПОЛЬЗОВАТЕЛЯ)
' ===============================================================================
Public Function ActivateProduct(key As String) As Boolean
    Dim parts() As String
    Dim rawKey As String
    Dim p1 As String, p2 As String, p3 As String, p4 As String
    Dim expectedCheck As String
    Dim dateVal As Long
    Dim expiryDate As Date
    
    ' 1. Форматирование
    key = UCase(Trim(key))
    ' Убираем лишние пробелы, если есть
    key = Replace(key, " ", "")
    
    parts = Split(key, "-")
    
    ' Проверка формата (должно быть 4 группы)
    If UBound(parts) <> 3 Then
        MsgBox "Неверный формат ключа. Должно быть 4 группы символов через тире.", vbExclamation
        ActivateProduct = False
        Exit Function
    End If
    
    p1 = parts(0)
    p2 = parts(1)
    p3 = parts(2)
    p4 = parts(3)
    
    ' 2. Проверка целостности (Контрольная сумма)
    expectedCheck = CalculateChecksum(p1 & p2 & p3)
    
    If p4 <> expectedCheck Then
        MsgBox "Ключ недействителен (ошибка контрольной суммы).", vbCritical
        ActivateProduct = False
        Exit Function
    End If
    
    ' 3. Расшифровка даты
    On Error Resume Next
    dateVal = CLng("&H" & p2) ' Из Hex в число
    dateVal = dateVal Xor CLng("&H" & p1) ' Обратный XOR с солью
    dateVal = dateVal + MAGIC_SEED ' Убираем смещение
    expiryDate = CDate(dateVal)
    
    If Err.number <> 0 Then
        MsgBox "Ключ поврежден.", vbCritical
        ActivateProduct = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' 4. Проверка срока действия самого ключа
    If expiryDate < Date Then
        MsgBox "Срок действия этого ключа истёк: " & Format(expiryDate, "dd.mm.yyyy"), vbExclamation
        ActivateProduct = False
        Exit Function
    End If
    
    ' 5. Успешная активация - Сохраняем в файл
    SaveLicenseState expiryDate
    
    MsgBox "Программа успешно активирована!" & vbCrLf & _
           "Лицензия действует до: " & Format(expiryDate, "dd.mm.yyyy"), vbInformation
           
    ActivateProduct = True
End Function

' ===============================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ===============================================================================

' Хэширование строки для контрольной суммы (возвращает 4 символа Hex)
Private Function CalculateChecksum(inputStr As String) As String
    Dim i As Long
    Dim hash As Long
    Dim fullStr As String
    
    ' Добавляем "Соль" приложения, чтобы нельзя было создать генератор, не зная константы
    fullStr = inputStr & SALT_KEY
    hash = 0
    
    For i = 1 To Len(fullStr)
        hash = (hash * 31 + Asc(Mid(fullStr, i, 1))) Mod 65535
    Next i
    
    CalculateChecksum = Right("0000" & Hex(hash), 4)
End Function

' Сохранение статуса в скрытые имена (Named Ranges)
' Это надежнее реестра, так как лицензия "путешествует" вместе с файлом
Private Sub SaveLicenseState(expDate As Date)
    On Error Resume Next
    Dim dateStr As String
    dateStr = CStr(CLng(expDate))
    
    ' Сохраняем дату
    ThisWorkbook.Names.Add Name:=NAME_LICENSE_DATA, RefersTo:="=" & dateStr, Visible:=False
    
    ' Сохраняем подпись (чтобы дату нельзя было просто поменять в диспетчере имен)
    Dim sign As String
    sign = CalculateChecksum(dateStr)
    ThisWorkbook.Names.Add Name:=NAME_LICENSE_SIGN, RefersTo:="=""" & sign & """", Visible:=False
End Sub

' Получение статуса лицензии
' 0 - Лицензия активна
' 1 - Лицензия истекла или отсутствует
Public Function GetLicenseStatus() As Integer
    On Error Resume Next
    Dim licDateRaw As String
    Dim storedSign As String
    Dim actualSign As String
    Dim expDate As Date
    
    ' Читаем дату
    licDateRaw = ThisWorkbook.Names(NAME_LICENSE_DATA).RefersTo
    licDateRaw = Replace(licDateRaw, "=", "")
    
    ' Читаем подпись
    storedSign = ThisWorkbook.Names(NAME_LICENSE_SIGN).RefersTo
    storedSign = Replace(Replace(storedSign, "=", ""), """", "")
    
    If licDateRaw = "" Then
        GetLicenseStatus = 1 ' Нет лицензии
        Exit Function
    End If
    
    ' Проверяем подпись (защита от взлома через изменение имени)
    actualSign = CalculateChecksum(licDateRaw)
    
    If storedSign <> actualSign Then
        GetLicenseStatus = 1 ' Данные повреждены или подделаны
        Exit Function
    End If
    
    ' Проверяем дату
    expDate = CDate(CLng(licDateRaw))
    If expDate >= Date Then
        GetLicenseStatus = 0 ' Активна
    Else
        GetLicenseStatus = 1 ' Истекла
    End If
End Function

Public Function GetLicenseExpiryDateStr() As String
    On Error Resume Next
    Dim v As String
    v = ThisWorkbook.Names(NAME_LICENSE_DATA).RefersTo
    v = Replace(v, "=", "")
    If IsNumeric(v) Then
        GetLicenseExpiryDateStr = Format(CDate(CLng(v)), "dd.mm.yyyy")
    Else
        GetLicenseExpiryDateStr = "Не активировано"
    End If
End Function

