Attribute VB_Name = "mdlPaymentValidation"
' ===============================================================================
' Module mdlPaymentValidation
' Version: 1.0.0
' Date: 14.02.2026
' Description: Validation of allowances without periods
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

' Column index constants for sheet "Выплаты_Без_Периодов"
Public Const COL_NUMBER As Long = 1          ' A
Public Const COL_PAYMENT_TYPE As Long = 2    ' B
Public Const COL_FIO As Long = 3             ' C
Public Const COL_LICHNIY_NOMER As Long = 4   ' D
Public Const COL_AMOUNT As Long = 5          ' E
Public Const COL_FOUNDATION As Long = 6      ' F
Public Const COL_PACKAGE_ID As Long = 7      ' G
Public Const COL_PACKAGE_MODE As Long = 8    ' H
Public Const COL_PARAMETER As Long = 9       ' I
Public Const COL_SHARED_FOUNDATION As Long = 10 ' J
Public Const COL_GROUP_EXPORT As Long = 11   ' K
Public Const COL_NOTE As Long = 12           ' L
Public Const COL_STATUS As Long = 13         ' M
Public Const COL_SOURCE_ENROLLMENT_ID As Long = 14 ' N

Private Const PAYMENT_STATUS_OK As String = "OK"
Private Const PAYMENT_STATUS_WARNING_KEY As String = "payments.validation.status.warning"
Private Const PAYMENT_STATUS_ERROR_KEY As String = "payments.validation.status.error"
Private Const PAYMENT_STATUS_DUPLICATE_KEY As String = "payments.validation.status.duplicate"
Private Const PAYMENT_STATUS_BLOCKED_KEY As String = "payments.validation.status.blocked"

Private Const PAYMENT_ELIGIBILITY_RULE_PARAM_REQUIRED As String = "PARAM_REQUIRED"
Private Const PAYMENT_ELIGIBILITY_RULE_POSITION_KEYWORDS As String = "POSITION_KEYWORDS"
Private Const PAYMENT_ELIGIBILITY_RULE_FOUNDATION_KEYWORDS As String = "FOUNDATION_KEYWORDS"
Private Const PAYMENT_ELIGIBILITY_SEVERITY_WARNING As String = "WARNING"
Private Const PAYMENT_ELIGIBILITY_SEVERITY_BLOCKED As String = "BLOCKED"

Private Function PaymentStatusWarning() As String
    PaymentStatusWarning = t(PAYMENT_STATUS_WARNING_KEY, "Warning")
End Function

Private Function PaymentStatusError() As String
    PaymentStatusError = t(PAYMENT_STATUS_ERROR_KEY, "Error")
End Function

Private Function PaymentStatusDuplicate() As String
    PaymentStatusDuplicate = t(PAYMENT_STATUS_DUPLICATE_KEY, "Package duplicate")
End Function

Private Function PaymentStatusBlocked() As String
    PaymentStatusBlocked = t(PAYMENT_STATUS_BLOCKED_KEY, "Blocked")
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Main function for validating all allowances
' =============================================
Public Sub ValidatePaymentsWithoutPeriods(Optional ByVal isSilent As Boolean = False)
    On Error GoTo ErrorHandler
    
    Dim wsPayments As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim errorCount As Long
    Dim warningCount As Long
    Dim reportText As String
    Dim isValid As Boolean
    Dim paymentType As String
    Dim duplicateKeys As Object
    Dim duplicateKey As String
    Dim packageSections As Object
    Dim warningText As String
    Dim statusText As String
    Dim rowHasError As Boolean
    
    Application.ScreenUpdating = False
    Application.StatusBar = t("payments.validation.status.running", "Validating allowances...")
    
    ' Search for sheet "Выплаты_Без_Периодов"
    Set wsPayments = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
            Set wsPayments = ws
            Exit For
        End If
    Next ws
    
    If wsPayments Is Nothing Then
        If Not isSilent Then MsgBox tf("payments.validation.error.sheet_missing", "Sheet ""{sheet}"" was not found.", "{sheet}", mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS), vbCritical, t("common.error", "Error")
        GoTo CleanUp
    End If

    mdlPaymentPackageSupport.EnsurePaymentsSheetEnhancements wsPayments
    
    lastRow = wsPayments.Cells(wsPayments.Rows.count, COL_LICHNIY_NOMER).End(xlUp).Row
    
    If lastRow < 2 Then
        If Not isSilent Then MsgBox tf("payments.validation.info.no_data", "Sheet ""{sheet}"" has no rows to validate.", "{sheet}", mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS), vbInformation, t("payments.validation.title.info", "Information")
        GoTo CleanUp
    End If
    
    errorCount = 0
    warningCount = 0
    Set duplicateKeys = CreateObject("Scripting.Dictionary")
    Set packageSections = BuildPackagePrimarySectionMap(wsPayments, lastRow)
    reportText = t("payments.validation.report.title", "====== ALLOWANCE VALIDATION REPORT ======") & vbCrLf & vbCrLf
    reportText = reportText & tf("payments.validation.report.date", "Check date: {date}", "{date}", Format(Now, "dd.mm.yyyy hh:mm:ss")) & vbCrLf
    reportText = reportText & tf("payments.validation.report.checked", "Checked records: {count}", "{count}", lastRow - 1) & vbCrLf & vbCrLf
    
    ' Check every row
    For i = 2 To lastRow
        Application.StatusBar = tf("payments.validation.status.row", "Checking row {row} of {total}", "{row}", i, "{total}", lastRow)
        wsPayments.Cells(i, COL_STATUS).Value = ""
        wsPayments.Cells(i, COL_STATUS).Interior.ColorIndex = xlNone
        warningText = ""
        statusText = PAYMENT_STATUS_OK
        rowHasError = False

        mdlPaymentPackageSupport.EnrichPaymentRow wsPayments, i
        
        paymentType = Trim(LCase(CStr(wsPayments.Cells(i, COL_PAYMENT_TYPE).value)))
        
        ' Call corresponding validation function
        Select Case paymentType
            Case "водители сдэ", "водители сде"
                isValid = ValidateDriverSDE(wsPayments, i)
            Case "экипаж"
                isValid = ValidateCrew(wsPayments, i)
            Case "физо"
                isValid = ValidateFIZO(wsPayments, i)
            Case "секретность"
                isValid = ValidateSecrecy(wsPayments, i)
            Case Else
                ' For unknown types - basic check
                isValid = ValidateBasic(wsPayments, i)
                If Not isValid Then
                    warningCount = warningCount + 1
                    reportText = reportText & tf("payments.validation.report.row_unknown_type", "Row {row}: Unknown payment type ""{type}""", "{row}", i, "{type}", paymentType) & vbCrLf
                End If
        End Select
        
        If Not isValid Then
            errorCount = errorCount + 1
            reportText = reportText & tf("payments.validation.report.row_type_error", "Row {row}: Validation failed for payment type ""{type}""", "{row}", i, "{type}", paymentType) & vbCrLf
            rowHasError = True
            statusText = PaymentStatusError()
        Else
            warningText = GetRowValidationWarning(wsPayments, i, packageSections)
            If warningText <> "" Then
                warningCount = warningCount + 1
                reportText = reportText & tf("payments.validation.report.row_message", "Row {row}: {message}", "{row}", i, "{message}", warningText) & vbCrLf
                statusText = PaymentStatusWarning()
            End If

            warningText = GetPaymentEligibilityIssue(wsPayments, i, rowHasError)
            If warningText <> "" Then
                reportText = reportText & tf("payments.validation.report.row_message", "Row {row}: {message}", "{row}", i, "{message}", warningText) & vbCrLf
                If rowHasError Then
                    errorCount = errorCount + 1
                    statusText = PaymentStatusBlocked()
                Else
                    warningCount = warningCount + 1
                    If statusText = PAYMENT_STATUS_OK Then statusText = PaymentStatusWarning()
                End If
            End If
        End If

        duplicateKey = BuildDuplicateKey(wsPayments, i)
        If duplicateKey <> "" Then
            If duplicateKeys.Exists(duplicateKey) Then
                errorCount = errorCount + 1
                reportText = reportText & tf("payments.validation.report.row_message", "Row {row}: {message}", "{row}", i, "{message}", t("payments.validation.warning.duplicate_employee_package", "Duplicate employee inside one package.")) & vbCrLf
                rowHasError = True
                statusText = PaymentStatusDuplicate()
            Else
                duplicateKeys.Add duplicateKey, True
            End If
        End If

        ApplyRowStatus wsPayments, i, statusText, rowHasError
    Next i
    
    ' Final report
    reportText = reportText & vbCrLf & t("payments.validation.report.total", "Total:") & vbCrLf
    reportText = reportText & tf("payments.validation.report.errors", "Errors: {count}", "{count}", errorCount) & vbCrLf
    reportText = reportText & tf("payments.validation.report.warnings", "Warnings: {count}", "{count}", warningCount) & vbCrLf
    
    If errorCount = 0 And warningCount = 0 Then
        reportText = reportText & vbCrLf & t("payments.validation.report.all_correct", "All data is correct.") & vbCrLf
        If Not isSilent Then MsgBox reportText, vbInformation, t("payments.validation.title.completed", "Validation completed")
    Else
        If Not isSilent Then MsgBox reportText, vbExclamation, t("payments.validation.title.completed", "Validation completed")
    End If
    
    GoTo CleanUp
    
ErrorHandler:
    If Not isSilent Then MsgBox tf("payments.validation.error.failed", "Allowance validation failed: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Error")
    
CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Basic validation (check mandatory fields)
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Private Function ValidateBasic(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fio As String
    Dim lichniyNomer As String
    Dim amount As String
    Dim foundation As String
    Dim paymentType As String
    Dim packageMode As String
    Dim packageId As String
    
    fio = Trim(CStr(ws.Cells(rowNum, COL_FIO).value))
    lichniyNomer = Trim(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).value))
    amount = mdlPaymentPackageSupport.GetEffectiveAmountFromSheet(ws, rowNum)
    foundation = mdlPaymentPackageSupport.GetEffectiveFoundationFromSheet(ws, rowNum)
    paymentType = Trim(CStr(ws.Cells(rowNum, COL_PAYMENT_TYPE).value))
    packageMode = mdlPaymentPackageSupport.NormalizeTextValue(CStr(ws.Cells(rowNum, COL_PACKAGE_MODE).value))
    packageId = Trim(CStr(ws.Cells(rowNum, COL_PACKAGE_ID).value))
    
    ' Check mandatory fields
    If paymentType = "" Or fio = "" Or lichniyNomer = "" Or amount = "" Or foundation = "" Then
        ValidateBasic = False
        Exit Function
    End If

    If packageMode = "LIST" And packageId = "" Then
        ValidateBasic = False
        Exit Function
    End If
    
    ValidateBasic = True
    Exit Function
    
ErrorHandler:
    ValidateBasic = False
End Function

Private Function BuildDuplicateKey(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim packageId As String
    Dim paymentType As String
    Dim personalNumber As String

    packageId = Trim$(CStr(ws.Cells(rowNum, COL_PACKAGE_ID).Value))
    If packageId = "" Then Exit Function

    paymentType = Trim$(LCase$(CStr(ws.Cells(rowNum, COL_PAYMENT_TYPE).Value)))
    personalNumber = Trim$(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).Value))

    If paymentType = "" Or personalNumber = "" Then Exit Function

    BuildDuplicateKey = LCase$(packageId) & "|" & paymentType & "|" & personalNumber
End Function

Private Sub ApplyRowStatus(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal statusText As String, ByVal rowHasError As Boolean)
    ws.Cells(rowNum, COL_STATUS).Value = statusText

    If rowHasError Then
        ws.Cells(rowNum, COL_STATUS).Interior.Color = RGB(255, 199, 206)
    ElseIf StrComp(statusText, PaymentStatusWarning(), vbTextCompare) = 0 Then
        ws.Cells(rowNum, COL_STATUS).Interior.Color = RGB(255, 235, 156)
    Else
        ws.Cells(rowNum, COL_STATUS).Interior.Color = RGB(198, 239, 206)
    End If
End Sub

Private Function GetRowValidationWarning(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal packageSections As Object) As String
    Dim paymentType As String
    Dim amountText As String
    Dim packageId As String
    Dim rowSection As String
    Dim primarySection As String
    Dim employeeExists As Boolean

    paymentType = Trim$(CStr(ws.Cells(rowNum, COL_PAYMENT_TYPE).Value))
    amountText = mdlPaymentPackageSupport.GetEffectiveAmountFromSheet(ws, rowNum)
    packageId = Trim$(CStr(ws.Cells(rowNum, COL_PACKAGE_ID).Value))

    If paymentType <> "" Then
        If Not PaymentTypeExists(paymentType) Then
            GetRowValidationWarning = AppendWarningText(GetRowValidationWarning, t("payments.validation.warning.type_not_in_reference", "Payment type is not in the reference."))
        End If
    End If

    employeeExists = EmployeeExistsInStaff(ws, rowNum)
    If Not employeeExists Then
        GetRowValidationWarning = AppendWarningText(GetRowValidationWarning, t("payments.validation.warning.employee_not_found", "Employee was not found on the Staff sheet."))
    End If

    If RequiresParameterValue(paymentType) Then
        If Trim$(CStr(ws.Cells(rowNum, COL_PARAMETER).Value)) = "" Then
            GetRowValidationWarning = AppendWarningText(GetRowValidationWarning, t("payments.validation.warning.parameter_missing", "Required payment parameter is missing."))
        End If
    End If

    If IsGroupedRow(ws, rowNum) Then
        If Trim$(CStr(ws.Cells(rowNum, COL_SHARED_FOUNDATION).Value)) = "" Then
            GetRowValidationWarning = AppendWarningText(GetRowValidationWarning, t("payments.validation.warning.shared_foundation_missing", "Shared basis is missing for package export."))
        End If
    End If

    If IsSuspiciousZeroAmount(amountText) Then
        GetRowValidationWarning = AppendWarningText(GetRowValidationWarning, t("payments.validation.warning.zero_amount", "Payment amount is empty or zero."))
    End If

    rowSection = GetEmployeeSectionForRow(ws, rowNum)
    If rowSection <> "" And packageId <> "" Then
        If packageSections.Exists(LCase$(packageId)) Then
            primarySection = Trim$(CStr(packageSections(LCase$(packageId))))
            If primarySection <> "" Then
                If StrComp(primarySection, rowSection, vbTextCompare) <> 0 Then
                    GetRowValidationWarning = AppendWarningText(GetRowValidationWarning, t("payments.validation.warning.other_section", "Employee belongs to another section inside the package."))
                End If
            End If
        End If
    End If
End Function

Private Function BuildPackagePrimarySectionMap(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim result As Object
    Dim counts As Object
    Dim rowNum As Long
    Dim packageId As String
    Dim sectionName As String
    Dim packageMap As Object
    Dim sectionCount As Long
    Dim packageKey As Variant
    Dim sectionKey As Variant
    Dim bestSection As String
    Dim bestCount As Long

    Set result = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")

    For rowNum = 2 To lastRow
        packageId = LCase$(Trim$(CStr(ws.Cells(rowNum, COL_PACKAGE_ID).Value)))
        If packageId <> "" Then
            sectionName = GetEmployeeSectionForRow(ws, rowNum)
            If sectionName <> "" Then
                If Not counts.Exists(packageId) Then
                    Set packageMap = CreateObject("Scripting.Dictionary")
                    counts.Add packageId, packageMap
                Else
                    Set packageMap = counts(packageId)
                End If

                sectionCount = 0
                If packageMap.Exists(sectionName) Then sectionCount = CLng(packageMap(sectionName))
                packageMap(sectionName) = sectionCount + 1
            End If
        End If
    Next rowNum

    For Each packageKey In counts.Keys
        Set packageMap = counts(packageKey)
        bestSection = ""
        bestCount = -1
        For Each sectionKey In packageMap.Keys
            sectionCount = CLng(packageMap(sectionKey))
            If sectionCount > bestCount Then
                bestCount = sectionCount
                bestSection = CStr(sectionKey)
            End If
        Next sectionKey
        If bestSection <> "" Then result(packageKey) = bestSection
    Next packageKey

    Set BuildPackagePrimarySectionMap = result
End Function

Private Function GetEmployeeSectionForRow(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim staffData As Object
    Dim personalNumber As String

    personalNumber = Trim$(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).Value))
    If personalNumber = "" Then Exit Function

    Set staffData = mdlHelper.GetStaffData(personalNumber, True)
    If staffData Is Nothing Then Exit Function
    If staffData.Count = 0 Then Exit Function

    GetEmployeeSectionForRow = Trim$(CStr(staffData("Часть")))
End Function

Private Function GetEmployeePositionForRow(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim staffData As Object
    Dim personalNumber As String

    personalNumber = Trim$(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).Value))
    If personalNumber = "" Then Exit Function

    Set staffData = mdlHelper.GetStaffData(personalNumber, True)
    If staffData Is Nothing Then Exit Function
    If staffData.Count = 0 Then Exit Function

    GetEmployeePositionForRow = Trim$(CStr(staffData("Штатная должность")))
End Function

Private Function GetEmployeeVusForRow(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim staffData As Object
    Dim personalNumber As String
    Dim vusKey As String

    personalNumber = Trim$(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).Value))
    If personalNumber = "" Then Exit Function

    Set staffData = mdlHelper.GetStaffData(personalNumber, True)
    If staffData Is Nothing Then Exit Function
    If staffData.Count = 0 Then Exit Function

    vusKey = mdlHelper.Ru(1042, 1059, 1057)
    If staffData.Exists(vusKey) Then
        GetEmployeeVusForRow = Trim$(CStr(staffData(vusKey)))
    End If
End Function

Private Function EmployeeExistsInStaff(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim staffData As Object
    Dim personalNumber As String

    personalNumber = Trim$(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).Value))
    If personalNumber = "" Then Exit Function

    Set staffData = mdlHelper.GetStaffData(personalNumber, True)
    If staffData Is Nothing Then Exit Function

    EmployeeExistsInStaff = (staffData.Count > 0)
End Function

Private Function PaymentTypeExists(ByVal paymentType As String) As Boolean
    Dim configDict As Object

    If Trim$(paymentType) = "" Then Exit Function

    Set configDict = mdlReferenceData.GetPaymentTypeConfig(paymentType)
    PaymentTypeExists = Not configDict Is Nothing
    If PaymentTypeExists Then PaymentTypeExists = (configDict.Count > 0)
End Function

Private Function RequiresParameterValue(ByVal paymentType As String) As Boolean
    Dim normalizedType As String

    normalizedType = LCase$(Trim$(paymentType))
    RequiresParameterValue = _
        normalizedType = LCase$(t("payments.type.class_qualification", mdlHelper.Ru(1050, 1083, 1072, 1089, 1089, 1085, 1072, 1103, 32, 1082, 1074, 1072, 1083, 1080, 1092, 1080, 1082, 1072, 1094, 1080, 1103)))
End Function

Private Function IsGroupedRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim modeValue As String
    Dim exportFlag As String

    modeValue = mdlPaymentPackageSupport.NormalizeTextValue(CStr(ws.Cells(rowNum, COL_PACKAGE_MODE).Value))
    exportFlag = mdlPaymentPackageSupport.NormalizeTextValue(CStr(ws.Cells(rowNum, COL_GROUP_EXPORT).Value))

    IsGroupedRow = (modeValue = "LIST" Or exportFlag = "ДА" Or exportFlag = "YES" Or exportFlag = "TRUE" Or exportFlag = "1")
End Function

Private Function IsSuspiciousZeroAmount(ByVal amountText As String) As Boolean
    Dim normalizedText As String
    Dim numericValue As Double

    normalizedText = Replace$(Replace$(Trim$(amountText), "%", ""), ",", ".")
    If normalizedText = "" Then
        IsSuspiciousZeroAmount = True
        Exit Function
    End If

    If IsNumeric(normalizedText) Then
        numericValue = CDbl(normalizedText)
        IsSuspiciousZeroAmount = (numericValue = 0)
    End If
End Function

Private Function AppendWarningText(ByVal currentText As String, ByVal newText As String) As String
    If Trim$(newText) = "" Then
        AppendWarningText = currentText
    ElseIf Trim$(currentText) = "" Then
        AppendWarningText = newText
    Else
        AppendWarningText = currentText & "; " & newText
    End If
End Function

Private Function GetPaymentEligibilityIssue(ByVal ws As Worksheet, ByVal rowNum As Long, ByRef isBlocked As Boolean) As String
    Dim config As mdlPaymentTypes.PaymentTypeConfig
    Dim paymentType As String
    Dim eligibilityRules As Variant
    Dim eligibilityRule As Variant
    Dim severityToken As String
    Dim issueText As String

    paymentType = Trim$(CStr(ws.Cells(rowNum, COL_PAYMENT_TYPE).Value))
    If paymentType = "" Then Exit Function

    config = mdlPaymentTypes.GetPaymentTypeConfig(paymentType)
    If Trim$(config.EligibilityRule) = "" Then Exit Function

    severityToken = NormalizePaymentEligibilitySeverity(config.EligibilitySeverity)
    eligibilityRules = Split(config.EligibilityRule, ";")

    For Each eligibilityRule In eligibilityRules
        issueText = EvaluateSinglePaymentEligibilityRule(ws, rowNum, paymentType, NormalizePaymentEligibilityRule(CStr(eligibilityRule)), config)
        If issueText <> "" Then
            GetPaymentEligibilityIssue = AppendWarningText(GetPaymentEligibilityIssue, issueText)
        End If
    Next eligibilityRule

    If GetPaymentEligibilityIssue <> "" Then
        isBlocked = (severityToken = PAYMENT_ELIGIBILITY_SEVERITY_BLOCKED)
    End If
End Function

Private Function EvaluateSinglePaymentEligibilityRule(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal paymentType As String, ByVal eligibilityRule As String, ByRef config As mdlPaymentTypes.PaymentTypeConfig) As String
    Dim employeePosition As String
    Dim foundationText As String

    Select Case eligibilityRule
        Case PAYMENT_ELIGIBILITY_RULE_PARAM_REQUIRED
            If Trim$(CStr(ws.Cells(rowNum, COL_PARAMETER).Value)) = "" Then
                EvaluateSinglePaymentEligibilityRule = tf("payments.validation.eligibility.param_required", "Eligibility condition failed for payment ""{type}"": required parameter is missing.", "{type}", paymentType)
            End If
        Case PAYMENT_ELIGIBILITY_RULE_POSITION_KEYWORDS
            employeePosition = GetEmployeePositionForRow(ws, rowNum)
            If employeePosition = "" Then
                EvaluateSinglePaymentEligibilityRule = tf("payments.validation.eligibility.position_missing", "Eligibility condition failed for payment ""{type}"": employee position could not be resolved.", "{type}", paymentType)
            ElseIf Not KeywordsMatchPaymentText(config.EligibilityPositionKeywords, employeePosition) Then
                EvaluateSinglePaymentEligibilityRule = tf("payments.validation.eligibility.position_mismatch", "Eligibility condition failed for payment ""{type}"": employee position does not match the rule.", "{type}", paymentType)
            End If
        Case PAYMENT_ELIGIBILITY_RULE_FOUNDATION_KEYWORDS
            foundationText = mdlPaymentPackageSupport.GetEffectiveFoundationFromSheet(ws, rowNum)
            If foundationText = "" Then
                EvaluateSinglePaymentEligibilityRule = tf("payments.validation.eligibility.foundation_missing", "Eligibility condition failed for payment ""{type}"": basis is missing.", "{type}", paymentType)
            ElseIf Not KeywordsMatchPaymentText(config.EligibilityFoundationKeywords, foundationText) Then
                EvaluateSinglePaymentEligibilityRule = tf("payments.validation.eligibility.foundation_mismatch", "Eligibility condition failed for payment ""{type}"": basis does not contain required eligibility markers.", "{type}", paymentType)
            End If
    End Select
End Function

Private Function NormalizePaymentEligibilityRule(ByVal ruleText As String) As String
    Select Case mdlPaymentPackageSupport.NormalizeTextValue(ruleText)
        Case PAYMENT_ELIGIBILITY_RULE_PARAM_REQUIRED
            NormalizePaymentEligibilityRule = PAYMENT_ELIGIBILITY_RULE_PARAM_REQUIRED
        Case PAYMENT_ELIGIBILITY_RULE_POSITION_KEYWORDS
            NormalizePaymentEligibilityRule = PAYMENT_ELIGIBILITY_RULE_POSITION_KEYWORDS
        Case PAYMENT_ELIGIBILITY_RULE_FOUNDATION_KEYWORDS
            NormalizePaymentEligibilityRule = PAYMENT_ELIGIBILITY_RULE_FOUNDATION_KEYWORDS
    End Select
End Function

Private Function KeywordsMatchPaymentText(ByVal keywordsText As String, ByVal sourceText As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim keyword As String
    Dim normalizedSource As String

    If Trim$(keywordsText) = "" Then
        KeywordsMatchPaymentText = True
        Exit Function
    End If

    normalizedSource = LCase$(Trim$(sourceText))
    parts = Split(keywordsText, ";")
    For i = LBound(parts) To UBound(parts)
        keyword = LCase$(Trim$(parts(i)))
        If keyword <> "" Then
            If InStr(1, normalizedSource, keyword, vbTextCompare) > 0 Then
                KeywordsMatchPaymentText = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function NormalizePaymentEligibilitySeverity(ByVal severityText As String) As String
    Select Case mdlPaymentPackageSupport.NormalizeTextValue(severityText)
        Case PAYMENT_ELIGIBILITY_SEVERITY_WARNING, "WARN", "INFO"
            NormalizePaymentEligibilitySeverity = PAYMENT_ELIGIBILITY_SEVERITY_WARNING
        Case Else
            NormalizePaymentEligibilitySeverity = PAYMENT_ELIGIBILITY_SEVERITY_BLOCKED
    End Select
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for Drivers CDE allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
' =============================================
' ИСПРАВЛЕННАЯ ВАЛИДАЦИЯ (Мягкая проверка)
' =============================================
' =============================================
' ИСПРАВЛЕННАЯ ВАЛИДАЦИЯ (Мягкая проверка для Водителей)
' =============================================
Public Function ValidateDriverSDE(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Базовые проверки (заполнены ли ФИО и номер)
    If Not ValidateBasic(ws, rowNum) Then ValidateDriverSDE = False: Exit Function
    
    ' 2. Проверка должности (через Штат) - опционально, если хотите строгую проверку должности
    ' Если нужно просто проверить текст основания, этот блок можно пропустить
    
    ' 3. Проверка текста основания (МЯГКАЯ)
    Dim foundation As String
    foundation = LCase(Trim(CStr(ws.Cells(rowNum, COL_FOUNDATION).value)))
    
    ' Мы ищем ХОТЯ БЫ ОДНО совпадение из списка ключевых слов
    Dim isValidDocs As Boolean
    isValidDocs = False
    
    ' Если есть "ваи" ИЛИ "ву" ИЛИ "приказ" ИЛИ "марка" - считаем верным
    If InStr(foundation, "ваи") > 0 Then isValidDocs = True
    If InStr(foundation, "ву") > 0 Then isValidDocs = True
    If InStr(foundation, "удостоверен") > 0 Then isValidDocs = True
    If InStr(foundation, "приказ") > 0 Then isValidDocs = True
    If InStr(foundation, "техник") > 0 Then isValidDocs = True
    
    ValidateDriverSDE = isValidDocs
    Exit Function
    
ErrorHandler:
    ValidateDriverSDE = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for Crew allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Public Function ValidateCrew(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lichniyNomer As String
    Dim staffData As Object
    Dim vus As String
    Dim Position As String
    
    ' Basic check
    If Not ValidateBasic(ws, rowNum) Then
        ValidateCrew = False
        Exit Function
    End If
    
    ' Get personal number
    lichniyNomer = Trim(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).value))
    
    ' Get data from staff
    Set staffData = mdlHelper.GetStaffData(lichniyNomer, True)
    If staffData.count = 0 Then
        ValidateCrew = False
        Exit Function
    End If
    
    Position = LCase(Trim(CStr(staffData("Штатная должность"))))
    vus = LCase$(GetEmployeeVusForRow(ws, rowNum))

    If vus <> "" And Position <> "" Then
        If mdlReferenceData.CheckVUSPositionPair(vus, Position) Then
            ValidateCrew = True
            Exit Function
        End If
    End If
    
    Dim crewKeywords As Variant
    crewKeywords = Array("командир", "механик", "наводчик", "оператор", "экипаж")
    
    Dim i As Long
    Dim hasCrewKeyword As Boolean
    hasCrewKeyword = False
    For i = LBound(crewKeywords) To UBound(crewKeywords)
        If InStr(Position, CStr(crewKeywords(i))) > 0 Then
            hasCrewKeyword = True
            Exit For
        End If
    Next i
    
    ValidateCrew = hasCrewKeyword
    Exit Function
    
ErrorHandler:
    ValidateCrew = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for FIZO allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Public Function ValidateFIZO(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim foundation As String
    Dim vedomostCount As Long
    Dim i As Long
    
    ' Basic check
    If Not ValidateBasic(ws, rowNum) Then
        ValidateFIZO = False
        Exit Function
    End If
    
    ' Get foundation
    foundation = LCase(Trim(CStr(ws.Cells(rowNum, COL_FOUNDATION).value)))
    
    ' Count occurrences of "vedomost"
    vedomostCount = 0
    i = 1
    Do While i <= Len(foundation)
        If Mid(foundation, i, 8) = "ведомость" Then
            vedomostCount = vedomostCount + 1
            i = i + 8
        Else
            i = i + 1
        End If
    Loop
    
    ' Must be at least 2 vedomosts
    ValidateFIZO = (vedomostCount >= 2)
    Exit Function
    
ErrorHandler:
    ValidateFIZO = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for Secrecy allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Public Function ValidateSecrecy(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim foundation As String
    Dim hasForm As Boolean, hasNumber As Boolean, hasDate As Boolean
    
    ' Базовая проверка на пустоту
    If Not ValidateBasic(ws, rowNum) Then
        ValidateSecrecy = False
        Exit Function
    End If
    
    foundation = Trim(CStr(ws.Cells(rowNum, COL_FOUNDATION).value))
    
    ' 1. Проверка Формы (ищем "форма 1", "форма 2", "форма 3" или "1 форма" и т.д.)
    ' Шаблон: слово "форма" рядом с цифрой 1, 2 или 3
    hasForm = mdlHelper.RegExpMatch(foundation, "(форма\s*[1-3]|[1-3]\s*форма)")
    
    ' 2. Проверка Номера (ищем "№ 123" или "номер 123")
    ' Шаблон: № или "номер" + пробелы + цифры/буквы
    hasNumber = mdlHelper.RegExpMatch(foundation, "(№|номер)\s*[\w\d-]+")
    
    ' 3. Проверка Даты (ищем формат ДД.ММ.ГГГГ или ДД.ММ.ГГ)
    ' Шаблон: цифры.цифры.цифры
    hasDate = mdlHelper.RegExpMatch(foundation, "\d{2}\.\d{2}\.\d{2,4}")
    
    ValidateSecrecy = (hasForm And hasNumber And hasDate)
    Exit Function
    
ErrorHandler:
    ValidateSecrecy = False
End Function
