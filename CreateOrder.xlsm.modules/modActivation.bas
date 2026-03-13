Attribute VB_Name = "modActivation"
' ===============================================================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Единая офлайн-система лицензирования (Trial + Personal/Corporate codes)
' @version 3.0.0
' ===============================================================================
Option Explicit

' --- ИНФОРМАЦИЯ О ПРОДУКТЕ (ДЛЯ ФОРМЫ FRMABOUT) ---
Public Const PRODUCT_NAME As String = "Формирователь приказов"
Public Const PRODUCT_VERSION As String = "2.5.0"
Public Const PRODUCT_AUTHOR As String = "Кержаев Евгений Алексеевич"
Public Const PRODUCT_EMAIL As String = "nachfin@vk.com"
Public Const PRODUCT_PHONE As String = "+7(989)906-88-91"
Public Const PRODUCT_COMPANY As String = "Отделение программирования 95 ФЭС МО РФ"
Public Const ACTIVATION_HINT As String = "Введите код (формат: XXXX-XXXX-XXXX-XXXX)"

' --- ГЛОБАЛЬНЫЕ НАСТРОЙКИ ЗАЩИТЫ ---
Private Const ADMIN_PASSWORD As String = "95FES_Admin"
Private Const SALT_KEY As String = "SECURE_TOKEN_2026"
Private Const MAGIC_SEED As Long = 1985

' --- АКТУАЛЬНЫЕ ИМЕНА СКРЫТЫХ ДИАПАЗОНОВ ---
Private Const NAME_LICENSE_CODE As String = "LicCode"
Private Const NAME_LAST_RUN As String = "LicLast"

' --- LEGACY ИМЕНА ДЛЯ МИГРАЦИИ СО СТАРОЙ СХЕМЫ ---
Private Const LEGACY_PERSONAL_DATA As String = "LicData"
Private Const LEGACY_PERSONAL_SIGN As String = "LicSign"
Private Const LEGACY_GLOBAL_DATA As String = "GlobalLimit"
Private Const LEGACY_GLOBAL_SIGN As String = "GlobalSign"

' --- ТИПЫ ЛИЦЕНЗИЙ ---
Private Const LICENSE_TYPE_PERSONAL As String = "PERSONAL"
Private Const LICENSE_TYPE_CORPORATE As String = "CORPORATE"
Private Const LICENSE_PREFIX_PERSONAL As String = "P"
Private Const LICENSE_PREFIX_CORPORATE As String = "C"
Private Const CORPORATE_VALIDATION_SALT As String = "CORPORATE"
Private Const LEGACY_GLOBAL_VALIDATION_SALT As String = "GLOBAL"

' ===============================================================================
' 0. ЕДИНЫЙ ИСТОЧНИК ИСТИНЫ ДЛЯ ОЗНАКОМИТЕЛЬНОГО ПЕРИОДА
' ===============================================================================
Private Function GetPublicExpiryDate() As Date
    GetPublicExpiryDate = DateSerial(2026, 9, 1) ' 1 Сентября 2026
End Function

' ===============================================================================
' 1. ИДЕНТИФИКАЦИЯ ОБОРУДОВАНИЯ (HWID)
' ===============================================================================
Public Function GetHardwareID() As String
    On Error Resume Next

    Dim fso As Object
    Dim d As Object
    Dim hexStr As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive("C:\")

    If Err.Number = 0 Then
        hexStr = Hex(d.SerialNumber)
        GetHardwareID = Right("00000000" & hexStr, 8)
    End If

    If GetHardwareID = "00000000" Or GetHardwareID = "" Then
        GetHardwareID = "NODEFAULT"
    End If

    On Error GoTo 0
End Function

' ===============================================================================
' 2. ГЕНЕРАЦИЯ КОДОВ ЛИЦЕНЗИИ
' ===============================================================================
Public Function GenerateLicenseKey(expiryDate As Date, targetHWID As String) As String
    GenerateLicenseKey = GenerateLicenseCode(expiryDate, LICENSE_TYPE_PERSONAL, targetHWID)
End Function

Public Function GenerateCorporateLicenseKey(expiryDate As Date) As String
    GenerateCorporateLicenseKey = GenerateLicenseCode(expiryDate, LICENSE_TYPE_CORPORATE)
End Function

Public Function GenerateLicenseCode(expiryDate As Date, _
                                    Optional licenseType As String = LICENSE_TYPE_PERSONAL, _
                                    Optional targetHWID As String = "") As String
    Dim normalizedType As String
    Dim validationSalt As String
    Dim p0_Header As String
    Dim p1_Date As String
    Dim p2_Noise As String
    Dim p3_Check As String
    Dim saltHex As String
    Dim saltValue As Long

    normalizedType = NormalizeLicenseType(licenseType)
    validationSalt = GetValidationSaltForType(normalizedType, targetHWID)
    If validationSalt = "" Then Exit Function

    Randomize

    saltHex = Right("000" & Hex(Int((4095 * Rnd) + 1)), 3)
    saltValue = CLng("&H" & saltHex)

    If normalizedType = LICENSE_TYPE_PERSONAL Then
        p0_Header = LICENSE_PREFIX_PERSONAL & saltHex
    Else
        p0_Header = LICENSE_PREFIX_CORPORATE & saltHex
    End If

    p1_Date = Right("0000" & Hex((CLng(expiryDate) - MAGIC_SEED) Xor saltValue), 4)
    p2_Noise = Right("0000" & Hex(Int((65535 * Rnd) + 1)), 4)
    p3_Check = CalculateHash(p0_Header & p1_Date & p2_Noise, validationSalt)

    GenerateLicenseCode = p0_Header & "-" & p1_Date & "-" & p2_Noise & "-" & p3_Check
End Function

' ===============================================================================
' 3. АКТИВАЦИЯ КОДОВ
' ===============================================================================
Public Function ActivatePersonal(key As String) As Boolean
    ActivatePersonal = ActivateLicenseCode(key)
End Function

Public Function ActivateLicenseCode(key As String) As Boolean
    Dim normalizedKey As String
    Dim licenseType As String
    Dim expiryDate As Date

    normalizedKey = NormalizeLicenseCode(key)
    If normalizedKey = "" Then
        ActivateLicenseCode = False
        Exit Function
    End If

    If Not TryParseLicenseCode(normalizedKey, licenseType, expiryDate) Then
        ActivateLicenseCode = False
        Exit Function
    End If

    If Date > expiryDate Then
        ActivateLicenseCode = False
        Exit Function
    End If

    SaveLicenseCode normalizedKey
    ClearLegacyLicenseData
    ActivateLicenseCode = True
End Function

' ===============================================================================
' 4. ГЛАВНАЯ ЛОГИКА ПРОВЕРКИ ЛИЦЕНЗИИ
' Возвращает:
' 0 - Активно (Персональный ключ)
' 3 - Активно (Корпоративный ключ)
' 4 - Активно (Ознакомительный период)
' 1 - Истекло / Нет лицензии
' 2 - Взлом времени
' ===============================================================================
Public Function GetLicenseStatus() As Integer
    Dim storedCode As String
    Dim licenseType As String
    Dim licenseExpiry As Date
    Dim publicExp As Date
    Dim lastRun As Date
    Dim lastRunStr As String
    Dim hasValidStoredCode As Boolean

    On Error Resume Next

    lastRunStr = ReadHidden(NAME_LAST_RUN)
    If lastRunStr <> "" Then
        lastRun = CDate(Val(lastRunStr))
        If Date < lastRun And lastRun > 0 Then
            GetLicenseStatus = 2
            Exit Function
        End If
    End If

    MigrateLegacyLicenseIfNeeded

    publicExp = GetPublicExpiryDate()
    storedCode = NormalizeLicenseCode(ReadHidden(NAME_LICENSE_CODE))

    If storedCode <> "" Then
        hasValidStoredCode = TryParseLicenseCode(storedCode, licenseType, licenseExpiry)
    End If

    UpdateLastRun

    If hasValidStoredCode And Date <= licenseExpiry Then
        If licenseType = LICENSE_TYPE_PERSONAL Then
            GetLicenseStatus = 0
        Else
            GetLicenseStatus = 3
        End If
        Exit Function
    End If

    If Date <= publicExp Then
        GetLicenseStatus = 4
    Else
        GetLicenseStatus = 1
    End If

    On Error GoTo 0
End Function

Public Function GetLicenseExpiryDateStr() As String
    Dim storedCode As String
    Dim licenseType As String
    Dim licenseExpiry As Date
    Dim effectiveExpiry As Date

    MigrateLegacyLicenseIfNeeded

    effectiveExpiry = GetPublicExpiryDate()
    storedCode = NormalizeLicenseCode(ReadHidden(NAME_LICENSE_CODE))

    If storedCode <> "" Then
        If TryParseLicenseCode(storedCode, licenseType, licenseExpiry) Then
            If Date <= licenseExpiry Then
                effectiveExpiry = licenseExpiry
            End If
        End If
    End If

    GetLicenseExpiryDateStr = Format(effectiveExpiry, "dd.mm.yyyy")
End Function

Public Function CheckLicenseAndPrompt() As Boolean
    Dim st As Integer

    st = GetLicenseStatus()

    If st = 0 Or st = 3 Or st = 4 Then
        CheckLicenseAndPrompt = True
    Else
        MsgBox "Срок действия доступа завершен." & vbCrLf & _
               "Для использования этой функции требуется действующий код активации.", _
               vbExclamation, "Требуется активация"
        frmAbout.Show

        st = GetLicenseStatus()
        CheckLicenseAndPrompt = (st = 0 Or st = 3 Or st = 4)
    End If
End Function

Public Sub EnsureLicenseOnOpen()
    Dim st As Integer

    st = GetLicenseStatus()
    If st = 1 Or st = 2 Then
        frmAbout.Show
    End If
End Sub

' ===============================================================================
' 5. СЛУЖЕБНЫЕ ПРОЦЕДУРЫ ДЛЯ ТЕСТИРОВАНИЯ И АДМИНИСТРАТОРА
' ===============================================================================
Public Sub ResetLicenseState()
    DeleteHidden NAME_LICENSE_CODE
    DeleteHidden NAME_LAST_RUN
    ClearLegacyLicenseData
End Sub

Public Sub AdminGeneratePersonalKeyUI()
    AdminGenerateKeyUI LICENSE_TYPE_PERSONAL
End Sub

Public Sub AdminGenerateCorporateKeyUI()
    AdminGenerateKeyUI LICENSE_TYPE_CORPORATE
End Sub

Public Sub AdminGenerateKeyUI(Optional suggestedType As String = "")
    Dim pwd As String
    Dim licenseChoice As String
    Dim licenseType As String
    Dim targetHWID As String
    Dim expDateStr As String
    Dim generatedKey As String

    pwd = InputBox("Введите пароль администратора для доступа к генератору:", "Генератор ключей")
    If pwd <> ADMIN_PASSWORD Then
        If pwd <> "" Then MsgBox "Неверный пароль!", vbCritical, "Отказ в доступе"
        Exit Sub
    End If

    If suggestedType <> "" Then
        licenseType = NormalizeLicenseType(suggestedType)
    Else
        licenseChoice = InputBox("Введите тип лицензии:" & vbCrLf & _
                                 "P - персональная" & vbCrLf & _
                                 "C - корпоративная", _
                                 "Тип лицензии", "C")
        licenseType = NormalizeLicenseType(licenseChoice)
    End If

    If licenseType = "" Then
        MsgBox "Не удалось определить тип лицензии.", vbExclamation, "Генератор ключей"
        Exit Sub
    End If

    If licenseType = LICENSE_TYPE_PERSONAL Then
        targetHWID = InputBox("Введите HWID компьютера пользователя:" & vbCrLf & _
                              "(Оставьте текущий, если делаете ключ для себя)", _
                              "Ввод HWID", GetHardwareID())
        If targetHWID = "" Then Exit Sub
    End If

    expDateStr = InputBox("Введите дату окончания действия ключа (ДД.ММ.ГГГГ):", _
                          "Дата окончания", "31.12.2026")
    If expDateStr = "" Then Exit Sub

    If Not IsDate(expDateStr) Then
        MsgBox "Некорректный формат даты! Используйте формат ДД.ММ.ГГГГ", _
               vbExclamation, "Ошибка"
        Exit Sub
    End If

    generatedKey = GenerateLicenseCode(CDate(expDateStr), licenseType, targetHWID)
    If generatedKey = "" Then
        MsgBox "Не удалось сформировать код лицензии.", vbCritical, "Ошибка"
        Exit Sub
    End If

    On Error Resume Next
    CreateObject("WScript.Shell").Run "cmd.exe /c echo | set /p=" & generatedKey & " | clip", 0, True
    On Error GoTo 0

    MsgBox "Ключ успешно сгенерирован:" & vbCrLf & vbCrLf & _
           generatedKey & vbCrLf & vbCrLf & _
           "Тип: " & GetLicenseTypeCaption(licenseType) & vbCrLf & _
           "Действует до: " & Format(CDate(expDateStr), "dd.mm.yyyy") & vbCrLf & vbCrLf & _
           "(Ключ уже скопирован в буфер обмена)", _
           vbInformation, "Успех"
End Sub

' ===============================================================================
' 6. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ===============================================================================
Private Sub MigrateLegacyLicenseIfNeeded()
    Dim legacyExpiry As Date
    Dim migratedCode As String

    If NormalizeLicenseCode(ReadHidden(NAME_LICENSE_CODE)) <> "" Then Exit Sub

    legacyExpiry = GetDateFromHidden(LEGACY_PERSONAL_DATA, LEGACY_PERSONAL_SIGN, True)
    If legacyExpiry > 0 Then
        migratedCode = GenerateLicenseCode(legacyExpiry, LICENSE_TYPE_PERSONAL, GetHardwareID())
        If migratedCode <> "" Then
            SaveLicenseCode migratedCode
            ClearLegacyLicenseData
        End If
        Exit Sub
    End If

    legacyExpiry = GetDateFromHidden(LEGACY_GLOBAL_DATA, LEGACY_GLOBAL_SIGN, False)
    If legacyExpiry > 0 Then
        migratedCode = GenerateLicenseCode(legacyExpiry, LICENSE_TYPE_CORPORATE)
        If migratedCode <> "" Then
            SaveLicenseCode migratedCode
            ClearLegacyLicenseData
        End If
    End If
End Sub

Private Function TryParseLicenseCode(ByVal key As String, _
                                     ByRef licenseType As String, _
                                     ByRef expiryDate As Date) As Boolean
    Dim p() As String
    Dim normalizedKey As String
    Dim typePrefix As String
    Dim saltValue As Long
    Dim decodedDate As Long
    Dim validationSalt As String

    normalizedKey = NormalizeLicenseCode(key)
    If normalizedKey = "" Then Exit Function

    p = Split(normalizedKey, "-")
    If UBound(p) <> 3 Then Exit Function

    typePrefix = Left$(p(0), 1)
    licenseType = LicenseTypeFromPrefix(typePrefix)
    If licenseType = "" Then Exit Function

    validationSalt = GetValidationSaltForType(licenseType)
    If validationSalt = "" Then Exit Function

    If p(3) <> CalculateHash(p(0) & p(1) & p(2), validationSalt) Then Exit Function

    On Error Resume Next
    saltValue = CLng("&H" & Right$(p(0), 3))
    decodedDate = CLng("&H" & p(1)) Xor saltValue
    expiryDate = CDate(decodedDate + MAGIC_SEED)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    TryParseLicenseCode = True
End Function

Private Function NormalizeLicenseCode(ByVal rawCode As String) As String
    Dim cleanText As String
    Dim formattedText As String
    Dim i As Long

    cleanText = UCase$(Trim$(rawCode))
    cleanText = Replace(cleanText, "-", "")
    cleanText = Replace(cleanText, " ", "")

    If Len(cleanText) <> 16 Then Exit Function

    For i = 1 To Len(cleanText)
        formattedText = formattedText & Mid$(cleanText, i, 1)
        If (i Mod 4 = 0) And i < Len(cleanText) Then
            formattedText = formattedText & "-"
        End If
    Next i

    NormalizeLicenseCode = formattedText
End Function

Private Function NormalizeLicenseType(ByVal rawType As String) As String
    Dim normalized As String

    normalized = UCase$(Trim$(rawType))

    Select Case normalized
        Case LICENSE_TYPE_PERSONAL, LICENSE_PREFIX_PERSONAL
            NormalizeLicenseType = LICENSE_TYPE_PERSONAL
        Case LICENSE_TYPE_CORPORATE, LICENSE_PREFIX_CORPORATE
            NormalizeLicenseType = LICENSE_TYPE_CORPORATE
    End Select
End Function

Private Function LicenseTypeFromPrefix(ByVal prefix As String) As String
    Select Case UCase$(Trim$(prefix))
        Case LICENSE_PREFIX_PERSONAL
            LicenseTypeFromPrefix = LICENSE_TYPE_PERSONAL
        Case LICENSE_PREFIX_CORPORATE
            LicenseTypeFromPrefix = LICENSE_TYPE_CORPORATE
    End Select
End Function

Private Function GetValidationSaltForType(ByVal licenseType As String, _
                                          Optional ByVal hwidOverride As String = "") As String
    Dim normalizedType As String
    Dim resolvedHWID As String

    normalizedType = NormalizeLicenseType(licenseType)

    Select Case normalizedType
        Case LICENSE_TYPE_PERSONAL
            resolvedHWID = UCase$(Trim$(hwidOverride))
            If resolvedHWID = "" Then resolvedHWID = GetHardwareID()
            If resolvedHWID = "" Or resolvedHWID = "NODEFAULT" Then Exit Function
            GetValidationSaltForType = resolvedHWID

        Case LICENSE_TYPE_CORPORATE
            GetValidationSaltForType = CORPORATE_VALIDATION_SALT
    End Select
End Function

Private Function GetLicenseTypeCaption(ByVal licenseType As String) As String
    Select Case NormalizeLicenseType(licenseType)
        Case LICENSE_TYPE_PERSONAL
            GetLicenseTypeCaption = "Персональная"
        Case LICENSE_TYPE_CORPORATE
            GetLicenseTypeCaption = "Корпоративная"
        Case Else
            GetLicenseTypeCaption = "Неизвестная"
    End Select
End Function

Private Function GetDateFromHidden(ByVal dataName As String, _
                                   ByVal signName As String, _
                                   ByVal useHWID As Boolean) As Date
    Dim dRaw As String
    Dim sRaw As String
    Dim expectedSign As String

    dRaw = ReadHidden(dataName)
    sRaw = ReadHidden(signName)

    If dRaw = "" Then
        GetDateFromHidden = 0
        Exit Function
    End If

    If useHWID Then
        expectedSign = CalculateHash(dRaw, GetHardwareID())
    Else
        expectedSign = CalculateHash(dRaw, LEGACY_GLOBAL_VALIDATION_SALT)
    End If

    If sRaw = expectedSign Then
        On Error Resume Next
        GetDateFromHidden = CDate(CLng(dRaw))
        On Error GoTo 0
    Else
        GetDateFromHidden = 0
    End If
End Function

Private Function CalculateHash(ByVal rawStr As String, ByVal saltExtra As String) As String
    Dim i As Long
    Dim hash As Long
    Dim fullStr As String

    fullStr = rawStr & saltExtra & SALT_KEY
    hash = 0

    For i = 1 To Len(fullStr)
        hash = (hash * 31 + Asc(Mid$(fullStr, i, 1))) Mod 65535
    Next i

    CalculateHash = Right$("0000" & Hex(hash), 4)
End Function

Private Sub SaveLicenseCode(ByVal licenseCode As String)
    WriteHidden NAME_LICENSE_CODE, NormalizeLicenseCode(licenseCode)
End Sub

Private Sub UpdateLastRun()
    WriteHidden NAME_LAST_RUN, CStr(CLng(Date))
End Sub

Private Sub ClearLegacyLicenseData()
    DeleteHidden LEGACY_PERSONAL_DATA
    DeleteHidden LEGACY_PERSONAL_SIGN
    DeleteHidden LEGACY_GLOBAL_DATA
    DeleteHidden LEGACY_GLOBAL_SIGN
End Sub

Private Sub DeleteHidden(ByVal nName As String)
    On Error Resume Next
    ThisWorkbook.Names(nName).Delete
    On Error GoTo 0
End Sub

Private Sub WriteHidden(ByVal nName As String, ByVal nValue As String)
    On Error Resume Next
    ThisWorkbook.Names(nName).Delete
    ThisWorkbook.Names.Add Name:=nName, RefersTo:="=""" & nValue & """", Visible:=False
    On Error GoTo 0
End Sub

Private Function ReadHidden(ByVal nName As String) As String
    On Error Resume Next

    Dim v As String

    v = ThisWorkbook.Names(nName).RefersTo
    If Err.Number = 0 Then
        v = Replace(v, "=", "")
        v = Replace(v, """", "")
        ReadHidden = v
    Else
        ReadHidden = ""
    End If

    On Error GoTo 0
End Function
