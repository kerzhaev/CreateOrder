Attribute VB_Name = "modActivation"
' ===============================================================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Единая офлайн-система лицензирования (Trial + Personal/Corporate codes)
' @version 3.0.0
' ===============================================================================
Option Explicit

' --- ИНФОРМАЦИЯ О ПРОДУКТЕ (ДЛЯ ФОРМЫ FRMABOUT) ---
Public Const PRODUCT_VERSION As String = "2.5.0"
Public Const PRODUCT_EMAIL As String = "nachfin@vk.com"
Public Const PRODUCT_PHONE As String = "+7(989)906-88-91"

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
Private Const ACTIVATION_REQUEST_HEADER As String = "CreateOrder Activation Request"
Private Const ACTIVATION_RESPONSE_HEADER As String = "CreateOrder Activation Response"
Private Const ACTIVATION_FILE_FORMAT_VERSION As String = "1"

' ===============================================================================
' 0. ЕДИНЫЙ ИСТОЧНИК ИСТИНЫ ДЛЯ ОЗНАКОМИТЕЛЬНОГО ПЕРИОДА
' ===============================================================================
Private Function GetPublicExpiryDate() As Date
    GetPublicExpiryDate = DateSerial(2026, 12, 31) ' 31 Декабря 2026
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

Public Function GetProductNameText() As String
    GetProductNameText = t("product.name", "Формирователь приказов")
End Function

Public Function GetProductAuthorText() As String
    GetProductAuthorText = t("product.author", "Кержаев Евгений Алексеевич")
End Function

Public Function GetProductCompanyText() As String
    GetProductCompanyText = t("product.company", "Отделение программирования 95 ФЭС МО РФ")
End Function

Public Function GetActivationHintText() As String
    GetActivationHintText = t("product.activation_hint", "Введите код (формат: XXXX-XXXX-XXXX-XXXX)")
End Function

Private Function TryParseDateExact(ByVal rawText As String, ByRef parsedDate As Date) As Boolean
    Dim cleanText As String
    Dim parts() As String
    Dim dayValue As Integer
    Dim monthValue As Integer
    Dim yearValue As Integer

    cleanText = Trim$(rawText)
    If cleanText = "" Then Exit Function

    parts = Split(cleanText, ".")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then Exit Function

    dayValue = CInt(parts(0))
    monthValue = CInt(parts(1))
    yearValue = CInt(parts(2))

    If yearValue < 100 Then yearValue = yearValue + 2000
    If dayValue < 1 Or dayValue > 31 Then Exit Function
    If monthValue < 1 Or monthValue > 12 Then Exit Function
    If yearValue < 2000 Or yearValue > 2099 Then Exit Function

    On Error GoTo ParseError
    parsedDate = DateSerial(yearValue, monthValue, dayValue)
    If Day(parsedDate) <> dayValue Or Month(parsedDate) <> monthValue Or Year(parsedDate) <> yearValue Then
        Exit Function
    End If

    TryParseDateExact = True
    Exit Function

ParseError:
    TryParseDateExact = False
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
    Dim activationMessage As String

    ActivateLicenseCode = TryActivateLicenseCodeDetailed(key, activationMessage)
End Function

Public Function TryActivateLicenseCodeDetailed(ByVal key As String, _
                                               ByRef resultMessage As String) As Boolean
    Dim normalizedKey As String
    Dim licenseType As String
    Dim expiryDate As Date
    Dim parseState As String

    normalizedKey = NormalizeLicenseCode(key)
    If normalizedKey = "" Then
        resultMessage = t("license.error.invalid_format", "Некорректный формат ключа. Используйте формат XXXX-XXXX-XXXX-XXXX.")
        TryActivateLicenseCodeDetailed = False
        Exit Function
    End If

    If Not TryParseLicenseCodeDetailed(normalizedKey, licenseType, expiryDate, parseState) Then
        resultMessage = GetActivationParseErrorMessage(parseState, licenseType)
        TryActivateLicenseCodeDetailed = False
        Exit Function
    End If

    If Date > expiryDate Then
        resultMessage = tf("license.error.expired", _
                           "Срок действия этого ключа истек {date}.", _
                           "{date}", Format$(expiryDate, "dd.mm.yyyy"))
        TryActivateLicenseCodeDetailed = False
        Exit Function
    End If

    SaveLicenseCode normalizedKey
    UpdateLastRun
    ClearLegacyLicenseData
    resultMessage = tf("license.success.activated", _
                       "Ключ успешно активирован. Лицензия действует до {date}.", _
                       "{date}", Format$(expiryDate, "dd.mm.yyyy"))
    TryActivateLicenseCodeDetailed = True
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

Public Function GetLicenseStatusCaption(Optional ByVal statusCode As Integer = -1) As String
    If statusCode < 0 Then statusCode = GetLicenseStatus()

    Select Case statusCode
        Case 0
            GetLicenseStatusCaption = t("license.status.personal_active", "Персональная лицензия активна")
        Case 3
            GetLicenseStatusCaption = t("license.status.corporate_active", "Корпоративная лицензия активна")
        Case 4
            GetLicenseStatusCaption = t("license.status.trial_active", "Ознакомительный период активен")
        Case 2
            GetLicenseStatusCaption = t("license.status.clock_lock", "Защитная блокировка по дате")
        Case Else
            GetLicenseStatusCaption = t("license.status.inactive", "Лицензия не активна")
    End Select
End Function

Public Function GetLicenseStatusDetailsText() As String
    Dim statusCode As Integer
    Dim storedCode As String
    Dim licenseType As String
    Dim licenseExpiry As Date
    Dim storedKeyLine As String
    Dim lastRunStr As String
    Dim publicExpiry As Date
    Dim details As String
    Dim publicExpiryLine As String

    statusCode = GetLicenseStatus()
    publicExpiry = GetPublicExpiryDate()
    storedCode = NormalizeLicenseCode(ReadHidden(NAME_LICENSE_CODE))
    lastRunStr = ReadHidden(NAME_LAST_RUN)

    If statusCode = 0 Or statusCode = 3 Then
        publicExpiryLine = tf("license.details.public_expiry_ignored", _
                              "Ознакомительный период до: {date} (не влияет при активной лицензии)", _
                              "{date}", Format$(publicExpiry, "dd.mm.yyyy"))
    Else
        publicExpiryLine = tf("license.details.public_expiry", _
                              "Ознакомительный период до: {date}", _
                              "{date}", Format$(publicExpiry, "dd.mm.yyyy"))
    End If

    details = tf("license.details.header", "Состояние лицензии: {status}", "{status}", GetLicenseStatusCaption(statusCode)) & vbCrLf & _
              tf("license.details.status_code", "Код состояния: {code}", "{code}", CStr(statusCode)) & vbCrLf & _
              tf("license.details.system_date", "Текущая дата системы: {date}", "{date}", Format$(Date, "dd.mm.yyyy")) & vbCrLf & _
              tf("license.details.system_datetime", "Текущее время системы: {datetime}", "{datetime}", Format$(Now, "dd.mm.yyyy HH:nn:ss")) & vbCrLf & _
              tf("license.details.hwid", "HWID: {hwid}", "{hwid}", GetHardwareID()) & vbCrLf & _
              tf("license.details.version", "Версия программы: {version}", "{version}", PRODUCT_VERSION) & vbCrLf & _
              tf("license.details.workbook_name", "Файл программы: {name}", "{name}", ThisWorkbook.Name) & vbCrLf & _
              publicExpiryLine & vbCrLf

    If lastRunStr <> "" Then
        details = details & tf("license.details.last_run", _
                               "Последний зафиксированный запуск: {date}", _
                               "{date}", Format$(CLng(lastRunStr), "dd.mm.yyyy")) & vbCrLf
    Else
        details = details & t("license.details.last_run_none", "Последний зафиксированный запуск: нет данных") & vbCrLf
    End If

    If storedCode <> "" Then
        If TryParseLicenseCode(storedCode, licenseType, licenseExpiry) Then
            storedKeyLine = tf("license.details.stored_key_ok", _
                               "Сохраненный ключ: {type}, до {date}, маска {mask}", _
                               "{type}", GetLicenseTypeCaption(licenseType), _
                               "{date}", Format$(licenseExpiry, "dd.mm.yyyy"), _
                               "{mask}", MaskLicenseCode(storedCode))
        Else
            storedKeyLine = t("license.details.stored_key_bad", "Сохраненный ключ: обнаружен, но не распознан")
        End If
    Else
        storedKeyLine = t("license.details.stored_key_none", "Сохраненный ключ: отсутствует")
    End If

    details = details & storedKeyLine & vbCrLf & _
              tf("license.details.recommendation", "Рекомендация: {text}", _
                 "{text}", GetLicenseRecommendation(statusCode))

    GetLicenseStatusDetailsText = details
End Function

Public Function ExportActivationRequestPackage() As Boolean
    Dim packageText As String
    Dim defaultName As String

    packageText = BuildActivationRequestText()
    defaultName = "CreateOrder-activation-request-" & _
                  Format$(Now, "yyyymmdd_HHnnss") & "-" & GetHardwareID() & ".txt"

    ExportActivationRequestPackage = SaveTextPackage(packageText, defaultName, _
                                                     t("license.export.file_filter", "Файлы активации (*.txt), *.txt"))
End Function

Public Function ImportActivationResponsePackage(Optional ByRef resultMessage As String) As Boolean
    Dim filePath As Variant
    Dim rawText As String
    Dim activationCode As String

    filePath = Application.GetOpenFilename(t("license.export.file_filter", "Файлы активации (*.txt), *.txt"), , _
                                           t("license.export.choose_response", "Выберите файл лицензии"))
    If VarType(filePath) = vbBoolean Then
        resultMessage = t("license.export.file_cancelled", "Выбор файла отменен.")
        Exit Function
    End If

    rawText = ReadTextFile(CStr(filePath))
    If Trim$(rawText) = "" Then
        resultMessage = t("license.export.file_read_error", "Не удалось прочитать файл лицензии.")
        Exit Function
    End If

    activationCode = ExtractActivationCodeFromPackage(rawText)
    If activationCode = "" Then
        resultMessage = t("license.export.code_missing", "В файле не найден код активации.")
        Exit Function
    End If

    ImportActivationResponsePackage = TryActivateLicenseCodeDetailed(activationCode, resultMessage)
End Function

Public Sub ShowLicenseServiceMenu()
    Dim answer As VbMsgBoxResult
    Dim importMessage As String

    answer = MsgBox(t("license.service.menu", _
                    "Да - подготовить файл запроса на лицензию" & vbCrLf & _
                    "Нет - загрузить готовый файл лицензии" & vbCrLf & _
                    "Отмена - показать подробное состояние лицензии"), _
                    vbYesNoCancel + vbQuestion, t("license.caption.service", "Сервис лицензии"))

    Select Case answer
        Case vbYes
            If ExportActivationRequestPackage() Then
                MsgBox t("license.export.saved", "Файл запроса на лицензию успешно сохранен."), _
                       vbInformation, t("license.caption.service", "Сервис лицензии")
            End If

        Case vbNo
            If ImportActivationResponsePackage(importMessage) Then
                MsgBox importMessage, vbInformation, t("license.caption.service", "Сервис лицензии")
            ElseIf importMessage <> "" Then
                MsgBox importMessage, vbExclamation, t("license.caption.service", "Сервис лицензии")
            End If

        Case Else
            MsgBox GetLicenseStatusDetailsText(), vbInformation, t("license.caption.state", "Состояние лицензии")
    End Select
End Sub

Public Sub ExportActivationRequestUI()
    If ExportActivationRequestPackage() Then
        MsgBox t("license.export.saved", "Файл запроса на лицензию успешно сохранен."), _
               vbInformation, t("license.caption.service", "Сервис лицензии")
    End If
End Sub

Public Sub ImportActivationResponseUI()
    Dim importMessage As String

    If ImportActivationResponsePackage(importMessage) Then
        MsgBox importMessage, vbInformation, t("license.caption.service", "Сервис лицензии")
    ElseIf importMessage <> "" Then
        MsgBox importMessage, vbExclamation, t("license.caption.service", "Сервис лицензии")
    End If
End Sub

Public Sub ShowLicenseStatusUI()
    MsgBox GetLicenseStatusDetailsText(), vbInformation, t("license.caption.state", "Состояние лицензии")
End Sub

Public Function CheckLicenseAndPrompt() As Boolean
    Dim st As Integer

    st = GetLicenseStatus()

    If st = 0 Or st = 3 Or st = 4 Then
        CheckLicenseAndPrompt = True
    Else
        MsgBox t("license.prompt.expired", _
               "Срок действия доступа завершен." & vbCrLf & "Для использования этой функции требуется действующий код активации."), _
               vbExclamation, t("license.caption.required", "Требуется активация")
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
    Dim expiryDate As Date
    Dim generatedKey As String

    pwd = InputBox(t("license.admin.password_prompt", "Введите пароль администратора для доступа к генератору:"), _
                   t("license.caption.generator", "Генератор ключей"))
    If pwd <> ADMIN_PASSWORD Then
        If pwd <> "" Then MsgBox t("license.admin.password_invalid", "Неверный пароль!"), _
                                 vbCritical, t("license.caption.access_denied", "Отказ в доступе")
        Exit Sub
    End If

    If suggestedType <> "" Then
        licenseType = NormalizeLicenseType(suggestedType)
    Else
        licenseChoice = InputBox(t("license.admin.type_prompt", _
                                 "Введите тип лицензии:" & vbCrLf & "P - персональная" & vbCrLf & "C - корпоративная"), _
                                 t("license.admin.type_title", "Тип лицензии"), "C")
        licenseType = NormalizeLicenseType(licenseChoice)
    End If

    If licenseType = "" Then
        MsgBox t("license.admin.type_unknown", "Не удалось определить тип лицензии."), _
               vbExclamation, t("license.caption.generator", "Генератор ключей")
        Exit Sub
    End If

    If licenseType = LICENSE_TYPE_PERSONAL Then
        targetHWID = InputBox(t("license.admin.hwid_prompt", _
                              "Введите HWID компьютера пользователя:" & vbCrLf & "(Оставьте текущий, если делаете ключ для себя)"), _
                              t("license.admin.hwid_title", "Ввод HWID"), GetHardwareID())
        If targetHWID = "" Then Exit Sub
    End If

    expDateStr = InputBox(t("license.admin.expiry_prompt", "Введите дату окончания действия ключа (ДД.ММ.ГГГГ):"), _
                          t("license.admin.expiry_title", "Дата окончания"), "31.12.2026")
    If expDateStr = "" Then Exit Sub

    If Not TryParseDateExact(expDateStr, expiryDate) Then
        MsgBox t("license.admin.expiry_invalid", "Некорректный формат даты! Используйте формат ДД.ММ.ГГГГ"), _
               vbExclamation, t("common.error", "Ошибка")
        Exit Sub
    End If

    generatedKey = GenerateLicenseCode(expiryDate, licenseType, targetHWID)
    If generatedKey = "" Then
        MsgBox t("license.admin.generate_failed", "Не удалось сформировать код лицензии."), _
               vbCritical, t("common.error", "Ошибка")
        Exit Sub
    End If

    On Error Resume Next
    CreateObject("WScript.Shell").Run "cmd.exe /c echo | set /p=" & generatedKey & " | clip", 0, True
    On Error GoTo 0

    MsgBox tf("license.admin.generate_success", _
              "Ключ успешно сгенерирован:" & vbCrLf & vbCrLf & "{key}" & vbCrLf & vbCrLf & "Тип: {type}" & vbCrLf & "Действует до: {date}" & vbCrLf & vbCrLf & "(Ключ уже скопирован в буфер обмена)", _
              "{key}", generatedKey, _
              "{type}", GetLicenseTypeCaption(licenseType), _
              "{date}", Format(expiryDate, "dd.mm.yyyy")), _
           vbInformation, t("license.caption.success", "Успех")
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
    Dim parseState As String

    TryParseLicenseCode = TryParseLicenseCodeDetailed(key, licenseType, expiryDate, parseState)
End Function

Private Function TryParseLicenseCodeDetailed(ByVal key As String, _
                                             ByRef licenseType As String, _
                                             ByRef expiryDate As Date, _
                                             Optional ByRef parseState As String = "") As Boolean
    Dim p() As String
    Dim normalizedKey As String
    Dim typePrefix As String
    Dim saltValue As Long
    Dim decodedDate As Long
    Dim validationSalt As String

    normalizedKey = NormalizeLicenseCode(key)
    If normalizedKey = "" Then
        parseState = "FORMAT"
        Exit Function
    End If

    p = Split(normalizedKey, "-")
    If UBound(p) <> 3 Then
        parseState = "FORMAT"
        Exit Function
    End If

    typePrefix = Left$(p(0), 1)
    licenseType = LicenseTypeFromPrefix(typePrefix)
    If licenseType = "" Then
        parseState = "TYPE"
        Exit Function
    End If

    validationSalt = GetValidationSaltForType(licenseType)
    If validationSalt = "" Then
        parseState = "HWID"
        Exit Function
    End If

    If p(3) <> CalculateHash(p(0) & p(1) & p(2), validationSalt) Then
        If licenseType = LICENSE_TYPE_PERSONAL Then
            parseState = "HWID_OR_CODE"
        Else
            parseState = "CHECKSUM"
        End If
        Exit Function
    End If

    On Error Resume Next
    saltValue = CLng("&H" & Right$(p(0), 3))
    decodedDate = CLng("&H" & p(1)) Xor saltValue
    expiryDate = CDate(decodedDate + MAGIC_SEED)
    If Err.Number <> 0 Then
        On Error GoTo 0
        parseState = "DATE"
        Exit Function
    End If
    On Error GoTo 0

    parseState = "OK"
    TryParseLicenseCodeDetailed = True
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
            GetLicenseTypeCaption = t("license.type.personal", "Персональная")
        Case LICENSE_TYPE_CORPORATE
            GetLicenseTypeCaption = t("license.type.corporate", "Корпоративная")
        Case Else
            GetLicenseTypeCaption = t("license.type.unknown", "Неизвестная")
    End Select
End Function

Private Function GetActivationParseErrorMessage(ByVal parseState As String, _
                                                ByVal licenseType As String) As String
    Select Case UCase$(parseState)
        Case "FORMAT"
            GetActivationParseErrorMessage = t("license.error.invalid_format", "Некорректный формат ключа. Используйте формат XXXX-XXXX-XXXX-XXXX.")
        Case "TYPE"
            GetActivationParseErrorMessage = t("license.parse.type", "Не удалось определить тип лицензии по введенному ключу.")
        Case "DATE"
            GetActivationParseErrorMessage = t("license.parse.date", "Ключ поврежден: не удалось прочитать дату окончания.")
        Case "HWID"
            GetActivationParseErrorMessage = t("license.parse.hwid", "Не удалось определить HWID этого компьютера для проверки персонального ключа.")
        Case "HWID_OR_CODE"
            GetActivationParseErrorMessage = t("license.parse.hwid_or_code", "Персональный ключ не подходит для этого ПК или был введен с ошибкой.")
        Case "CHECKSUM"
            If NormalizeLicenseType(licenseType) = LICENSE_TYPE_CORPORATE Then
                GetActivationParseErrorMessage = t("license.parse.corporate_checksum", "Корпоративный ключ не распознан. Проверьте правильность введенного кода.")
            Else
                GetActivationParseErrorMessage = t("license.parse.generic", "Ключ не распознан. Проверьте правильность введенного кода.")
            End If
        Case Else
            GetActivationParseErrorMessage = t("license.parse.generic", "Ключ не распознан. Проверьте правильность введенного кода.")
    End Select
End Function

Private Function BuildActivationRequestText() As String
    Dim statusCode As Integer
    Dim storedCode As String
    Dim requestId As String
    Dim details As String

    statusCode = GetLicenseStatus()
    storedCode = NormalizeLicenseCode(ReadHidden(NAME_LICENSE_CODE))
    requestId = Format$(Now, "yyyymmdd_HHnnss") & "_" & GetHardwareID()
    details = Replace$(GetLicenseStatusDetailsText(), vbCrLf, vbLf)

    BuildActivationRequestText = ACTIVATION_REQUEST_HEADER & vbCrLf & _
                                 "FormatVersion: " & ACTIVATION_FILE_FORMAT_VERSION & vbCrLf & _
                                 "RequestID: " & requestId & vbCrLf & _
                                 "CreatedAt: " & Format$(Now, "dd.mm.yyyy HH:nn:ss") & vbCrLf & _
                                 "Product: " & GetProductNameText() & vbCrLf & _
                                 "ProductVersion: " & PRODUCT_VERSION & vbCrLf & _
                                 "WorkbookName: " & ThisWorkbook.Name & vbCrLf & _
                                 "WorkbookPath: " & ThisWorkbook.FullName & vbCrLf & _
                                 "ComputerName: " & Environ$("COMPUTERNAME") & vbCrLf & _
                                 "UserName: " & Environ$("USERNAME") & vbCrLf & _
                                 "SystemDate: " & Format$(Date, "dd.mm.yyyy") & vbCrLf & _
                                 "SystemDateTime: " & Format$(Now, "dd.mm.yyyy HH:nn:ss") & vbCrLf & _
                                 "HWID: " & GetHardwareID() & vbCrLf & _
                                 "LicenseStatusCode: " & CStr(statusCode) & vbCrLf & _
                                 "LicenseStatusText: " & GetLicenseStatusCaption(statusCode) & vbCrLf & _
                                 "LicenseExpiry: " & GetLicenseExpiryDateStr() & vbCrLf & _
                                 "StoredLicenseMask: " & MaskLicenseCode(storedCode) & vbCrLf & _
                                 "RecommendedLicenseType: PERSONAL" & vbCrLf & _
                                 "Details: " & details & vbCrLf
End Function

Private Function GetLicenseRecommendation(ByVal statusCode As Integer) As String
    Select Case statusCode
        Case 0, 3
            GetLicenseRecommendation = t("license.recommendation.active", "Лицензия уже активна. При необходимости можно подготовить диагностический файл.")
        Case 4
            GetLicenseRecommendation = t("license.recommendation.trial", "Можно заранее подготовить файл запроса на персональную или корпоративную лицензию.")
        Case 2
            GetLicenseRecommendation = t("license.recommendation.clock_lock", "Проверьте корректность системной даты. При необходимости загрузите новый файл лицензии.")
        Case Else
            GetLicenseRecommendation = t("license.recommendation.inactive", "Подготовьте файл запроса на лицензию или загрузите готовый файл лицензии.")
    End Select
End Function

Private Function MaskLicenseCode(ByVal licenseCode As String) As String
    Dim normalizedCode As String

    normalizedCode = NormalizeLicenseCode(licenseCode)
    If normalizedCode = "" Then
        MaskLicenseCode = "нет"
    Else
        MaskLicenseCode = Left$(normalizedCode, 4) & "-****-****-" & Right$(normalizedCode, 4)
    End If
End Function

Private Function SaveTextPackage(ByVal packageText As String, _
                                 ByVal defaultFileName As String, _
                                 Optional ByVal fileFilter As String = "Text Files (*.txt), *.txt") As Boolean
    Dim targetPath As Variant
    Dim fileNumber As Integer

    If Trim$(packageText) = "" Then Exit Function

    targetPath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFileName, _
        FileFilter:=fileFilter)

    If VarType(targetPath) = vbBoolean Then Exit Function

    fileNumber = FreeFile
    Open CStr(targetPath) For Output As #fileNumber
    Print #fileNumber, packageText
    Close #fileNumber

    SaveTextPackage = True
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fileNumber As Integer

    On Error GoTo ReadError

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    ReadTextFile = Input$(LOF(fileNumber), #fileNumber)
    Close #fileNumber
    Exit Function

ReadError:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    ReadTextFile = ""
End Function

Private Function ExtractActivationCodeFromPackage(ByVal packageText As String) As String
    Dim normalizedText As String
    Dim lines() As String
    Dim i As Long
    Dim currentLine As String

    normalizedText = Replace$(packageText, vbCrLf, vbLf)
    normalizedText = Replace$(normalizedText, vbCr, vbLf)
    lines = Split(normalizedText, vbLf)

    For i = LBound(lines) To UBound(lines)
        currentLine = Trim$(lines(i))
        If UCase$(Left$(currentLine, 5)) = "CODE:" Then
            ExtractActivationCodeFromPackage = NormalizeLicenseCode(Trim$(Mid$(currentLine, 6)))
            If ExtractActivationCodeFromPackage <> "" Then Exit Function

            If i < UBound(lines) Then
                ExtractActivationCodeFromPackage = NormalizeLicenseCode(Trim$(lines(i + 1)))
                If ExtractActivationCodeFromPackage <> "" Then Exit Function
            End If
        ElseIf UCase$(Left$(currentLine, 15)) = "ACTIVATIONCODE:" Then
            ExtractActivationCodeFromPackage = NormalizeLicenseCode(Trim$(Mid$(currentLine, 16)))
            If ExtractActivationCodeFromPackage <> "" Then Exit Function
        End If
    Next i
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
