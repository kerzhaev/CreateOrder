Attribute VB_Name = "mdlUniversalPaymentExport"
' ===============================================================================
' Module mdlUniversalPaymentExport
' Version: 1.0.0
' Date: 14.02.2026
' Description: Universal export of allowances without periods to Word
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Main function for mass import of employees by numbers
' =============================================
Public Sub ImportEmployeesByNumbers()

    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub

    On Error GoTo ErrorHandler

    Dim wsPayments As Worksheet
    Dim selectedRange As Range
    Dim reportText As String

    Application.ScreenUpdating = False
    Application.StatusBar = t("payments.import.status.mass_add", "Mass employee import...")

    ' Check active sheet
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox tf("payments.import.error.wrong_sheet", "Open sheet ""{sheet}"".", "{sheet}", mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS), vbExclamation, t("common.error", "Error")
        GoTo CleanUp
    End If

    Set wsPayments = ActiveSheet

    ' Get selected range
    Set selectedRange = Selection
    If selectedRange Is Nothing Then
        MsgBox t("payments.import.error.select_numbers", "Select cells with personal numbers in column D."), vbExclamation, t("common.error", "Error")
        GoTo CleanUp
    End If

    ' Check if range is in column D
    If selectedRange.Column <> mdlPaymentValidation.COL_LICHNIY_NOMER Then
        MsgBox t("payments.import.error.select_column_d", "Selected range must be in column D (personal number)."), vbExclamation, t("common.error", "Error")
        GoTo CleanUp
    End If

    ' Process range
    reportText = ProcessSelectedRangeForImport(wsPayments, selectedRange)

    ' Show report
    MsgBox reportText, vbInformation, t("payments.import.title.completed", "Mass import completed")

    GoTo CleanUp

ErrorHandler:
    MsgBox tf("payments.import.error.failed", "Mass employee import failed: {error}", "{error}", Err.description), vbCritical, t("common.error", "Error")

CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Process selected range and fill employee data
' @param wsPayments As Worksheet - sheet "Выплаты_Без_Периодов"
' @param selectedRange As Range - selected range
' @return String - report on results
' =============================================
Public Function ProcessSelectedRangeForImport(wsPayments As Worksheet, selectedRange As Range) As String
    On Error GoTo ErrorHandler

    Dim cell As Range
    Dim numberValue As String
    Dim staffData As Object
    Dim foundCount As Long
    Dim notFoundCount As Long
    Dim notFoundList As String
    Dim reportText As String
    Dim lastRow As Long
    Dim i As Long

    foundCount = 0
    notFoundCount = 0
    notFoundList = ""

    Application.ScreenUpdating = False

    ' Process each cell in range
    For Each cell In selectedRange.Cells
        numberValue = Trim(CStr(cell.value))

        ' Skip empty cells
        If numberValue = "" Then
            GoTo NextCell
        End If

        ' Find employee by number (personal or table)
        Set staffData = mdlHelper.FindEmployeeByAnyNumber(numberValue)

        If staffData.count > 0 Then
            ' Employee found - fill data
            cell.value = CStr(staffData("Личный номер")) ' Update personal number (normalization)
            wsPayments.Cells(cell.Row, mdlPaymentValidation.COL_FIO).value = CStr(staffData("Лицо")) ' Fill FIO
            foundCount = foundCount + 1
        Else
            ' Not found - add to list
            notFoundCount = notFoundCount + 1
            If notFoundList <> "" Then
                notFoundList = notFoundList & ", "
            End If
            notFoundList = notFoundList & numberValue
        End If

NextCell:
    Next cell

    ' Automatic numbering in column A
    lastRow = wsPayments.Cells(wsPayments.Rows.count, mdlPaymentValidation.COL_LICHNIY_NOMER).End(xlUp).Row
    For i = 2 To lastRow
        If Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_NUMBER).value)) = "" Then
            wsPayments.Cells(i, mdlPaymentValidation.COL_NUMBER).value = i - 1
        End If
    Next i

    ' Generate report
    reportText = t("payments.import.report.title", "Mass import results:") & vbCrLf & vbCrLf
    reportText = reportText & tf("payments.import.report.found", "Found and added: {count}", "{count}", foundCount) & vbCrLf
    reportText = reportText & tf("payments.import.report.not_found", "Not found: {count}", "{count}", notFoundCount)

    If notFoundCount > 0 And Len(notFoundList) < 200 Then
        reportText = reportText & vbCrLf & tf("payments.import.report.not_found_numbers", "Numbers not found: {numbers}", "{numbers}", notFoundList)
    ElseIf notFoundCount > 0 Then
        reportText = reportText & vbCrLf & t("payments.import.report.not_found_too_long", "(Missing number list is too long to display)")
    End If

    ProcessSelectedRangeForImport = reportText
    Exit Function

ErrorHandler:
    ProcessSelectedRangeForImport = tf("payments.import.error.process_range", "Range processing failed: {error}", "{error}", Err.description)
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Convert PaymentWithoutPeriod to Dictionary (for storing in Collection)
' @param payment As PaymentWithoutPeriod - payment data
' @return Object (Dictionary) - dictionary with payment data
' =============================================
Private Function PaymentToDictionary(ByRef payment As PaymentWithoutPeriod) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict("fio") = payment.fio
    dict("lichniyNomer") = payment.lichniyNomer
    dict("Rank") = payment.Rank
    dict("Position") = payment.Position
    dict("VoinskayaChast") = payment.VoinskayaChast
    dict("paymentType") = payment.paymentType
    dict("amount") = payment.amount
    dict("foundation") = payment.foundation
    dict("packageId") = payment.packageId
    dict("packageMode") = payment.packageMode
    dict("parameterValue") = payment.parameterValue
    dict("sharedFoundation") = payment.sharedFoundation
    dict("groupExportFlag") = payment.groupExportFlag
    dict("noteText") = payment.noteText
    dict("statusText") = payment.statusText
    dict("sourceEnrollmentId") = payment.sourceEnrollmentId
    Set PaymentToDictionary = dict
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Convert Dictionary back to PaymentWithoutPeriod
' @param dict As Object (Dictionary) - dictionary with payment data
' @return PaymentWithoutPeriod - payment data
' =============================================
Private Function DictionaryToPayment(ByRef dict As Object) As PaymentWithoutPeriod
    Dim payment As PaymentWithoutPeriod
    If dict.count > 0 Then
        payment.fio = CStr(dict("fio"))
        payment.lichniyNomer = CStr(dict("lichniyNomer"))
        payment.Rank = CStr(dict("Rank"))
        payment.Position = CStr(dict("Position"))
        payment.VoinskayaChast = CStr(dict("VoinskayaChast"))
        payment.paymentType = CStr(dict("paymentType"))
        payment.amount = CStr(dict("amount"))
        payment.foundation = CStr(dict("foundation"))
        payment.packageId = CStr(dict("packageId"))
        payment.packageMode = CStr(dict("packageMode"))
        payment.parameterValue = CStr(dict("parameterValue"))
        payment.sharedFoundation = CStr(dict("sharedFoundation"))
        payment.groupExportFlag = CStr(dict("groupExportFlag"))
        payment.noteText = CStr(dict("noteText"))
        payment.statusText = CStr(dict("statusText"))
        payment.sourceEnrollmentId = CStr(dict("sourceEnrollmentId"))
    End If
    DictionaryToPayment = payment
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Main function for exporting allowances
' =============================================
Public Sub ExportPaymentsWithoutPeriods()
    Call ExportPaymentsWithoutPeriodsCore(True)
End Sub

Public Function ExportPaymentsWithoutPeriodsCore(Optional ByVal showMessages As Boolean = True) As Long
    On Error GoTo ErrorHandler

    Dim payments As Collection
    Dim groupedPayments As Object
    Dim paymentGroupKey As Variant
    Dim paymentList As Collection
    Dim successCount As Long
    Dim payment As PaymentWithoutPeriod
    Dim paymentDict As Object

    Application.ScreenUpdating = False
    Application.StatusBar = t("payments.export.status.collecting", "Collecting allowance data...")

    mdlPaymentValidation.ValidatePaymentsWithoutPeriods True

    ' Collect all payment data
    Set payments = CollectPaymentsData(True)

    If payments.count = 0 Then
        If showMessages Then MsgBox tf("payments.export.error.no_ready_rows", "No rows ready for export on sheet ""{sheet}"".", "{sheet}", mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS), vbExclamation, t("payments.export.title", "Export")
        GoTo CleanUp
    End If

    Application.StatusBar = t("payments.export.status.grouping", "Grouping payment types...")

    ' Group by export package / type
    Set groupedPayments = GroupPaymentsByType(payments)

    Application.StatusBar = t("payments.export.status.generating", "Generating orders...")

    ' Generate orders for each payment type
    successCount = 0
    For Each paymentGroupKey In groupedPayments.keys
        Set paymentList = groupedPayments(paymentGroupKey)
        Set paymentDict = paymentList(1)
        payment = DictionaryToPayment(paymentDict)
        If GeneratePaymentOrder(CStr(payment.paymentType), paymentList, GetExportSuffix(payment), showMessages) Then
            successCount = successCount + 1
        End If
    Next paymentGroupKey

    ExportPaymentsWithoutPeriodsCore = successCount
    If showMessages Then
        MsgBox tf("payments.export.message.completed", "Export completed. Created orders: {success} of {total}", "{success}", successCount, "{total}", groupedPayments.count) & vbCrLf & _
               t("payments.export.message.saved_to_workbook_folder", "Files are saved to the workbook folder:") & vbCrLf & ThisWorkbook.Path, _
               vbInformation, t("payments.export.title", "Export")
    End If

    GoTo CleanUp

ErrorHandler:
    If showMessages Then MsgBox tf("payments.export.error.failed", "Allowance export failed: {error}", "{error}", Err.description), vbCritical, t("common.error", "Error")

CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Collect all payment data from sheet "Выплаты_Без_Периодов"
' @return Collection - collection of PaymentWithoutPeriod objects
' =============================================
Public Function CollectPaymentsData(Optional ByVal skipBlockedRows As Boolean = False) As Collection
    On Error GoTo ErrorHandler

    Dim wsPayments As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim result As Collection
    Dim payment As PaymentWithoutPeriod

    Set result = New Collection

    ' Find sheet "Выплаты_Без_Периодов"
    Set wsPayments = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
            Set wsPayments = ws
            Exit For
        End If
    Next ws

    If wsPayments Is Nothing Then
        Set CollectPaymentsData = result
        Exit Function
    End If

    lastRow = wsPayments.Cells(wsPayments.Rows.count, mdlPaymentValidation.COL_LICHNIY_NOMER).End(xlUp).Row

    If lastRow < 2 Then
        Set CollectPaymentsData = result
        Exit Function
    End If

    ' Collect data from each row
    For i = 2 To lastRow
        payment = BuildPaymentFromSheetRow(wsPayments, i)

        ' Skip empty rows
        If payment.lichniyNomer = "" Then
            GoTo nextRow
        End If

        If skipBlockedRows Then
            If IsPaymentRowBlockedForExport(payment.statusText) Then
                GoTo nextRow
            End If
        End If

        ' Convert UDT to Dictionary for storage in Collection
        result.Add PaymentToDictionary(payment)

nextRow:
    Next i

    Set CollectPaymentsData = result
    Exit Function

ErrorHandler:
    Set CollectPaymentsData = New Collection
End Function

Private Function IsPaymentRowBlockedForExport(ByVal statusText As String) As Boolean
    Dim normalizedStatus As String

    normalizedStatus = mdlPaymentPackageSupport.NormalizeTextValue(statusText)
    Select Case normalizedStatus
        Case "ОШИБКА", "ДУБЛИКАТ ПАКЕТА", "БЛОКИРОВАНО", "ERROR", "BLOCKED"
            IsPaymentRowBlockedForExport = True
    End Select
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Group payments by type
' @param payments As Collection - collection of payments
' @return Object (Dictionary) - dictionary where key = payment type, value = collection of payments
' =============================================
Public Function GroupPaymentsByType(ByVal payments As Collection) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim payment As PaymentWithoutPeriod
    Dim paymentGroupKey As String
    Dim paymentList As Collection

    Set result = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim paymentDict As Object
    For i = 1 To payments.count
        ' Extract Dictionary from Collection and convert to UDT
        Set paymentDict = payments(i)
        payment = DictionaryToPayment(paymentDict)
        paymentGroupKey = mdlPaymentPackageSupport.BuildExportGroupKey(payment)

        If paymentGroupKey = "" Then
            paymentGroupKey = "не_указан"
        End If

        If Not result.exists(paymentGroupKey) Then
            Set paymentList = New Collection
            result.Add paymentGroupKey, paymentList
        Else
            Set paymentList = result(paymentGroupKey)
        End If

        ' Add Dictionary back to paymentList
        paymentList.Add paymentDict
    Next i

    Set GroupPaymentsByType = result
    Exit Function

ErrorHandler:
    Set GroupPaymentsByType = CreateObject("Scripting.Dictionary")
End Function

Public Function BuildPaymentFromSheetRow(ByVal wsPayments As Worksheet, ByVal rowNum As Long) As PaymentWithoutPeriod
    Dim payment As PaymentWithoutPeriod
    Dim staffData As Object

    payment.fio = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_FIO).Value))
    payment.lichniyNomer = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_LICHNIY_NOMER).Value))
    payment.paymentType = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_PAYMENT_TYPE).Value))
    payment.amount = mdlPaymentPackageSupport.GetEffectiveAmountFromSheet(wsPayments, rowNum)
    payment.foundation = mdlPaymentPackageSupport.GetEffectiveFoundationFromSheet(wsPayments, rowNum)
    payment.packageId = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_PACKAGE_ID).Value))
    payment.packageMode = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_PACKAGE_MODE).Value))
    payment.parameterValue = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_PARAMETER).Value))
    payment.sharedFoundation = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value))
    payment.groupExportFlag = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_GROUP_EXPORT).Value))
    payment.noteText = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_NOTE).Value))
    payment.statusText = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_STATUS).Value))
    payment.sourceEnrollmentId = Trim$(CStr(wsPayments.Cells(rowNum, mdlPaymentValidation.COL_SOURCE_ENROLLMENT_ID).Value))

    Set staffData = Nothing
    If payment.lichniyNomer <> "" Then
        Set staffData = mdlHelper.GetStaffData(payment.lichniyNomer, True)
    End If

    If Not staffData Is Nothing Then
        If staffData.Count > 0 Then
            payment.Rank = CStr(staffData("Воинское звание"))
            payment.Position = CStr(staffData("Штатная должность"))
            payment.VoinskayaChast = mdlHelper.ExtractVoinskayaChast(CStr(staffData("Часть")))
            If payment.fio = "" Then payment.fio = CStr(staffData("Лицо"))
        Else
            payment.Rank = t("payments.export.fallback.rank_not_found", "Rank not found")
            payment.Position = t("payments.export.fallback.position_not_found", "Position not found")
            payment.VoinskayaChast = ""
        End If
    End If

    BuildPaymentFromSheetRow = payment
End Function

Public Function BuildPaymentPreviewTextFromRows(ByVal wsPayments As Worksheet, ByVal rowNumbers As Collection, Optional ByVal headerText As String = "") As String
    Dim previewText As String
    Dim groupedRows As Object
    Dim groupOrder As Collection
    Dim i As Long
    Dim groupKey As String
    Dim payment As PaymentWithoutPeriod
    Dim paymentList As Collection
    Dim orderKey As Variant

    If Trim$(headerText) <> "" Then
        previewText = headerText & vbCrLf & String(60, "=") & vbCrLf & vbCrLf
    End If

    Set groupedRows = CreateObject("Scripting.Dictionary")
    Set groupOrder = New Collection

    For i = 1 To rowNumbers.Count
        payment = BuildPaymentFromSheetRow(wsPayments, CLng(rowNumbers(i)))
        If payment.lichniyNomer <> "" Then
            groupKey = mdlPaymentPackageSupport.BuildExportGroupKey(payment)
            If groupKey = "" Then groupKey = "ROW|" & CStr(CLng(rowNumbers(i)))

            If Not groupedRows.Exists(groupKey) Then
                Set paymentList = New Collection
                groupedRows.Add groupKey, paymentList
                groupOrder.Add groupKey
            Else
                Set paymentList = groupedRows(groupKey)
            End If

            paymentList.Add PaymentToDictionary(payment)
        End If
    Next i

    For i = 1 To groupOrder.Count
        orderKey = groupOrder(i)
        Set paymentList = groupedRows(CStr(orderKey))
        previewText = previewText & BuildPaymentPreviewGroupText(paymentList)
        If i < groupOrder.Count Then
            previewText = previewText & vbCrLf & vbCrLf & String(60, "-") & vbCrLf & vbCrLf
        End If
    Next i

    BuildPaymentPreviewTextFromRows = previewText
End Function

Private Function BuildPaymentPreviewGroupText(ByVal payments As Collection) As String
    Dim previewText As String
    Dim firstPayment As PaymentWithoutPeriod
    Dim payment As PaymentWithoutPeriod
    Dim i As Long
    Dim commonFoundation As String
    Dim hasSharedFoundation As Boolean

    If payments.Count = 0 Then Exit Function

    firstPayment = DictionaryToPayment(payments(1))
    hasSharedFoundation = TryGetCommonFoundation(payments, commonFoundation)

    If mdlPaymentPackageSupport.ShouldUseGroupedExport(firstPayment) And payments.Count > 1 Then
        previewText = PreviewPaymentTypeLabel() & ": " & firstPayment.paymentType
        If Trim$(firstPayment.packageId) <> "" Then
            previewText = previewText & vbCrLf & PreviewPackageLabel() & ": " & Trim$(firstPayment.packageId)
        End If
        If hasSharedFoundation Then
            previewText = previewText & vbCrLf & PreviewFoundationLabel() & ": " & commonFoundation
        End If
        previewText = previewText & vbCrLf

        For i = 1 To payments.Count
            payment = DictionaryToPayment(payments(i))
            previewText = previewText & BuildPreviewEmployeeLine(payment, i, Not hasSharedFoundation)
            If i < payments.Count Then
                previewText = previewText & vbCrLf & vbCrLf
            End If
        Next i
    Else
        For i = 1 To payments.Count
            payment = DictionaryToPayment(payments(i))
            previewText = previewText & BuildPreviewEmployeeLine(payment, i, True)
            If i < payments.Count Then
                previewText = previewText & vbCrLf & vbCrLf
            End If
        Next i
    End If

    BuildPaymentPreviewGroupText = previewText
End Function

Private Function TryGetCommonFoundation(ByVal payments As Collection, ByRef commonFoundation As String) As Boolean
    Dim i As Long
    Dim payment As PaymentWithoutPeriod
    Dim foundationText As String

    If payments.Count = 0 Then Exit Function

    For i = 1 To payments.Count
        payment = DictionaryToPayment(payments(i))
        foundationText = Trim$(payment.foundation)
        If foundationText = "" Then Exit Function

        If commonFoundation = "" Then
            commonFoundation = foundationText
        ElseIf StrComp(commonFoundation, foundationText, vbTextCompare) <> 0 Then
            commonFoundation = ""
            Exit Function
        End If
    Next i

    TryGetCommonFoundation = (commonFoundation <> "")
End Function

Private Function BuildPreviewEmployeeLine(ByRef payment As PaymentWithoutPeriod, ByVal index As Long, ByVal includeFoundation As Boolean) As String
    Dim previewLine As String
    Dim foundationMarker As String
    Dim markerPosition As Long

    previewLine = FormatEmployeePaymentText(payment, index)
    If includeFoundation Then
        BuildPreviewEmployeeLine = previewLine
        Exit Function
    End If

    foundationMarker = vbCrLf & PreviewFoundationLabel() & ": "
    markerPosition = InStr(1, previewLine, foundationMarker, vbTextCompare)
    If markerPosition > 0 Then
        previewLine = Left$(previewLine, markerPosition - 1)
    End If

    BuildPreviewEmployeeLine = previewLine
End Function

Private Function BuildGroupedExportText(ByVal payments As Collection) As String
    Dim exportText As String
    Dim commonFoundation As String
    Dim hasSharedFoundation As Boolean
    Dim payment As PaymentWithoutPeriod
    Dim i As Long

    If payments.Count = 0 Then Exit Function

    hasSharedFoundation = TryGetCommonFoundation(payments, commonFoundation)

    For i = 1 To payments.Count
        payment = DictionaryToPayment(payments(i))
        exportText = exportText & BuildExportEmployeeLine(payment, i, Not hasSharedFoundation)
        If i < payments.Count Then
            exportText = exportText & vbCrLf & vbCrLf
        End If
    Next i

    If hasSharedFoundation Then
        exportText = exportText & vbCrLf & vbCrLf & FoundationLabel() & ": " & commonFoundation
    End If

    BuildGroupedExportText = exportText
End Function

Private Function BuildExportEmployeeLine(ByRef payment As PaymentWithoutPeriod, ByVal index As Long, ByVal includeFoundation As Boolean) As String
    Dim exportLine As String
    Dim foundationMarker As String
    Dim markerPosition As Long

    exportLine = FormatEmployeePaymentText(payment, index)
    If includeFoundation Then
        BuildExportEmployeeLine = exportLine
        Exit Function
    End If

    foundationMarker = vbCrLf & FoundationLabel() & ": "
    markerPosition = InStr(1, exportLine, foundationMarker, vbTextCompare)
    If markerPosition > 0 Then
        exportLine = Left$(exportLine, markerPosition - 1)
    End If

    BuildExportEmployeeLine = exportLine
End Function

Private Function FoundationLabel() As String
    FoundationLabel = t("payments.foundation.label", mdlHelper.Ru(1054, 1089, 1085, 1086, 1074, 1072, 1085, 1080, 1077))
End Function

Private Function PreviewFoundationLabel() As String
    PreviewFoundationLabel = FoundationLabel()
End Function

Private Function PreviewPackageLabel() As String
    PreviewPackageLabel = t("payments.preview.package", mdlHelper.Ru(1055, 1072, 1082, 1077, 1090))
End Function

Private Function PreviewPaymentTypeLabel() As String
    PreviewPaymentTypeLabel = t("payments.preview.payment_type", mdlHelper.Ru(1058, 1080, 1087, 32, 1074, 1099, 1087, 1083, 1072, 1090, 1099))
End Function

Private Function GetExportSuffix(ByRef payment As PaymentWithoutPeriod) As String
    Dim rawSuffix As String

    rawSuffix = Trim$(payment.packageId)
    If rawSuffix = "" Then rawSuffix = Trim$(payment.sourceEnrollmentId)
    If rawSuffix = "" Then Exit Function

    rawSuffix = Replace$(rawSuffix, "\", "_")
    rawSuffix = Replace$(rawSuffix, "/", "_")
    rawSuffix = Replace$(rawSuffix, ":", "_")
    rawSuffix = Replace$(rawSuffix, "*", "_")
    rawSuffix = Replace$(rawSuffix, "?", "_")
    rawSuffix = Replace$(rawSuffix, """", "_")
    rawSuffix = Replace$(rawSuffix, "<", "_")
    rawSuffix = Replace$(rawSuffix, ">", "_")
    rawSuffix = Replace$(rawSuffix, "|", "_")
    rawSuffix = Replace$(rawSuffix, " ", "_")

    GetExportSuffix = rawSuffix
End Function


' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Generate Word order for specific payment type
' @param paymentType As String - payment type
' @param payments As Collection - collection of payments of this type
' @return Boolean - True if order successfully created
' =============================================
Public Function GeneratePaymentOrder(ByVal paymentType As String, ByVal payments As Collection, Optional ByVal fileSuffix As String = "", Optional ByVal showMessages As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim wordApp As Object
    Dim doc As Object
    Dim templateDoc As Object
    Dim config As PaymentTypeConfig
    Dim templatePath As String
    Dim payment As PaymentWithoutPeriod
    Dim i As Long
    Dim fileName As String
    Dim savePath As String
    Dim wordWasNotRunning As Boolean
    Dim successCount As Long
    Dim paymentDict As Object
    Dim endRange As Object
    Dim isListTemplate As Boolean
    Dim listText As String
    Dim useGroupedBlock As Boolean

    ' Get payment type configuration
    config = mdlPaymentTypes.GetPaymentTypeConfig(paymentType)

    ' Get template path with priority
    templatePath = mdlPaymentTypes.GetTemplatePathWithFallback(config)

    If payments.Count > 0 Then
        Set paymentDict = payments(1)
        payment = DictionaryToPayment(paymentDict)
        useGroupedBlock = mdlPaymentPackageSupport.ShouldUseGroupedExport(payment) And payments.Count > 1
    End If

    Set wordApp = CreateObject("Word.Application")
    If wordApp Is Nothing Then
        Err.Raise vbObjectError + 713, "GeneratePaymentOrder", t("payments.export.error.word_instance", "Could not create a separate Word instance.")
    End If
    wordWasNotRunning = True
    On Error GoTo ErrorHandler

    wordApp.Visible = showMessages

    ' Determine if we are using a List Template
    isListTemplate = False
    If templatePath <> "" Then
        Set doc = wordApp.Documents.Add(templatePath)
        ' Проверяем наличие маркера списка
        With doc.content.Find
            .Text = "[СПИСОК_ВОЕННОСЛУЖАЩИХ]"
            If .Execute Then
                isListTemplate = True
            End If
        End With

        If Not isListTemplate Then
            doc.Close False

            If useGroupedBlock Then
                ' Для grouped-export без списочного маркера используем единый текстовый блок,
                ' иначе общее основание дублируется в универсальном шаблоне.
                Set doc = wordApp.Documents.Add
                With doc.Styles(1).Font
                    .Name = "Times New Roman"
                    .Size = 12
                End With
            Else
                ' Если это не списочный шаблон, используем старую логику постраничного копирования
                Set doc = wordApp.Documents.Add

                Set templateDoc = wordApp.Documents.Open(templatePath)
                templateDoc.content.Copy
                doc.content.Paste
                templateDoc.Close False
                Set templateDoc = Nothing
            End If
        End If
    Else
        Set doc = wordApp.Documents.Add
        With doc.Styles(1).Font
            .Name = "Times New Roman"
            .Size = 12
        End With
    End If

    successCount = 0

    If isListTemplate Then
        ' ==========================================
        ' НОВАЯ ЛОГИКА: ШАБЛОН СО СПИСКОМ
        ' ==========================================
        If useGroupedBlock Then
            listText = BuildGroupedExportText(payments)
        Else
            listText = ""
            For i = 1 To payments.count
                Set paymentDict = payments(i)
                payment = DictionaryToPayment(paymentDict)

                ' Используем нашу новую функцию форматирования
                Dim empText As String
                empText = FormatEmployeePaymentText(payment, i)

                ' Разделитель между военнослужащими (пустая строка)
                If i < payments.count Then
                    empText = empText & vbCrLf
                End If

                listText = listText & empText
            Next i
        End If

        ' Вставляем готовый список на место маркера, обходя ограничение в 255 символов
        Dim rngFind As Object
        Set rngFind = doc.content
        With rngFind.Find
            .ClearFormatting
            .Text = "[СПИСОК_ВОЕННОСЛУЖАЩИХ]"
            .Forward = True
            .Wrap = 0 ' wdFindStop
            If .Execute Then
                rngFind.Text = listText
            End If
        End With

        successCount = payments.count

    Else
        ' ==========================================
        ' СТАРАЯ ЛОГИКА: ИНДИВИДУАЛЬНЫЕ МАРКЕРЫ (ИЛИ БЕЗ ШАБЛОНА)
        ' ==========================================
        If useGroupedBlock And templatePath = "" Then
            If GenerateGroupedPaymentTextDirectly(doc, payments) Then
                successCount = payments.count
            End If
        Else
            For i = 1 To payments.count
                Set paymentDict = payments(i)
                payment = DictionaryToPayment(paymentDict)

                If templatePath <> "" Then
                    If i = 1 Then
                        If FillPaymentTemplate(doc, payment) Then
                            successCount = successCount + 1
                        End If
                    Else
                        Set templateDoc = wordApp.Documents.Open(templatePath)
                        If FillPaymentTemplate(templateDoc, payment) Then
                            templateDoc.content.Copy
                            Set endRange = doc.Range
                            endRange.Collapse Direction:=0
                            If i > 1 Then
                                endRange.InsertAfter vbCrLf & vbCrLf
                                endRange.Collapse Direction:=0
                            End If
                            endRange.Paste
                            successCount = successCount + 1
                        End If
                        templateDoc.Close False
                        Set templateDoc = Nothing
                    End If
                Else
                    If i > 1 Then
                        Set endRange = doc.Range
                        endRange.Collapse Direction:=0
                        endRange.InsertAfter vbCrLf & vbCrLf
                    End If

                    ' Передаем i для нумерации
                    If GeneratePaymentTextDirectly(doc, payment, i) Then
                        successCount = successCount + 1
                    End If
                End If
            Next i
        End If
    End If

    ' Сохранение
    Dim cleanTypeName As String
    cleanTypeName = Replace(Replace(Replace(paymentType, " ", "_"), "/", "_"), "\", "_")
    fileName = t("payments.export.filename.prefix", "Order") & "_" & cleanTypeName & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    If config.TypeCode <> "" Then
        fileName = t("payments.export.filename.prefix", "Order") & "_" & config.TypeCode & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    End If
    If Trim$(fileSuffix) <> "" Then
        fileName = Replace$(fileName, ".docx", "_" & Trim$(fileSuffix) & ".docx")
    End If
    savePath = ThisWorkbook.Path & "\" & fileName

    Call mdlHelper.SaveWordDocumentSafe(doc, savePath)
    If showMessages Then
        doc.Activate
        MsgBox tf("payments.export.message.order_created", "Created order with {success} records of {total}", "{success}", successCount, "{total}", payments.count), vbInformation, t("payments.export.title.completed", "Export completed")
    Else
        doc.Close False
        wordApp.Quit False
        Set doc = Nothing
        Set wordApp = Nothing
        wordWasNotRunning = False
    End If

    GeneratePaymentOrder = (successCount > 0)
    Exit Function

ErrorHandler:
    GeneratePaymentOrder = False
    If Not templateDoc Is Nothing Then templateDoc.Close False
    If Not doc Is Nothing Then doc.Close False
    If wordWasNotRunning And Not wordApp Is Nothing Then wordApp.Quit False
    MsgBox tf("payments.export.error.create_order", "Order creation failed: {error}", "{error}", Err.description), vbCritical, t("common.error", "Error")
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Fill Word template with payment data
' @param doc As Object - Word document
' @param payment As PaymentWithoutPeriod - payment data
' @return Boolean - True if successful
' =============================================
Public Function FillPaymentTemplate(ByVal doc As Object, ByRef payment As PaymentWithoutPeriod) As Boolean
    On Error GoTo ErrorHandler

    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ФИО]", payment.fio)
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ФИО_ИМЕНИТЕЛЬНЫЙ]", mdlHelper.SklonitFIO(payment.fio))
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ЗВАНИЕ]", payment.Rank)
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ЗВАНИЕ_СКЛОНЕННОЕ]", mdlHelper.SklonitZvanie(payment.Rank))
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ЛИЧНЫЙ_НОМЕР]", payment.lichniyNomer)
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ДОЛЖНОСТЬ]", payment.Position)
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ДОЛЖНОСТЬ_СКЛОНЕННАЯ]", mdlHelper.SklonitDolzhnost(payment.Position, payment.VoinskayaChast))
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[РАЗМЕР]", payment.amount)
    Call mdlWordTemplateSafe.ReplacePlaceholderText(doc, "[ОСНОВАНИЕ]", payment.foundation)

    FillPaymentTemplate = True
    Exit Function

ErrorHandler:
    FillPaymentTemplate = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Generate order text directly in Word without template
' @param doc As Object - Word document
' @param payment As PaymentWithoutPeriod - payment data
' @return Boolean - True if successful
' =============================================
' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Generate order text directly in Word without template
' =============================================
Public Function GeneratePaymentTextDirectly(ByVal doc As Object, ByRef payment As PaymentWithoutPeriod, ByVal index As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Object
    Dim textLine As String

    ' Используем общую функцию форматирования
    textLine = FormatEmployeePaymentText(payment, index) & vbCrLf

    ' Вставляем текст в документ
    Set rng = doc.Range
    rng.Collapse Direction:=0
    rng.Text = textLine
    rng.Font.Name = "Times New Roman"
    rng.Font.Size = 14

    GeneratePaymentTextDirectly = True
    Exit Function

ErrorHandler:
    GeneratePaymentTextDirectly = False
End Function

Public Function GenerateGroupedPaymentTextDirectly(ByVal doc As Object, ByVal payments As Collection) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Object
    Dim textBlock As String

    textBlock = BuildGroupedExportText(payments)
    If Trim$(textBlock) = "" Then Exit Function

    Set rng = doc.Range
    rng.Collapse Direction:=0
    rng.Text = textBlock & vbCrLf
    rng.Font.Name = "Times New Roman"
    rng.Font.Size = 14

    GenerateGroupedPaymentTextDirectly = True
    Exit Function

ErrorHandler:
    GenerateGroupedPaymentTextDirectly = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Формирует готовый текст для одного военнослужащего со всеми проверками
' =============================================
Private Function FormatEmployeePaymentText(ByRef payment As PaymentWithoutPeriod, ByVal index As Long) As String
    Dim cleanFIO As String, cleanRank As String, cleanPos As String, cleanVC As String, cleanFound As String

    ' 1. Очищаем все данные от случайных переносов строк (Alt+Enter из Excel)
    cleanFIO = Replace(Replace(payment.fio, vbCr, ""), vbLf, " ")
    cleanRank = Replace(Replace(payment.Rank, vbCr, ""), vbLf, " ")
    cleanPos = Replace(Replace(payment.Position, vbCr, ""), vbLf, " ")
    cleanVC = Replace(Replace(payment.VoinskayaChast, vbCr, ""), vbLf, " ")
    cleanFound = Replace(Replace(payment.foundation, vbCr, ""), vbLf, " ")

    ' Убираем двойные пробелы, если они появились после замены
    While InStr(cleanPos, "  ") > 0: cleanPos = Replace(cleanPos, "  ", " "): Wend
    While InStr(cleanFound, "  ") > 0: cleanFound = Replace(cleanFound, "  ", " "): Wend

    ' 2. Преобразуем размер (0,3 -> 30) независимым от системы способом
    Dim formattedAmount As String
    Dim numVal As Double

    formattedAmount = Trim(payment.amount)
    formattedAmount = Replace(formattedAmount, "%", "") ' Сначала убираем знак процента, если он был

    ' Функция Val понимает только точку, поэтому принудительно меняем запятую на точку
    Dim dotAmount As String
    dotAmount = Replace(formattedAmount, ",", ".")

    ' Если внутри действительно число
    If IsNumeric(dotAmount) Or IsNumeric(formattedAmount) Then
        numVal = val(dotAmount)
        ' Если число дробное (от 0.01 до 1), умножаем на 100
        If numVal > 0 And numVal <= 1 Then
            formattedAmount = CStr(numVal * 100)
        End If
    End If

    ' 3. Формируем текст БЕЗ лишних переносов строк
    Dim textLine As String
    textLine = index & ". " & mdlHelper.SklonitZvanie(cleanRank) & " " & _
               mdlHelper.SklonitFIO(cleanFIO) & ", " & tf("payments.export.word.personal_number", "personal number {number}", "{number}", payment.lichniyNomer) & ", " & _
               mdlHelper.SklonitDolzhnost(cleanPos, cleanVC)

    ' Добавляем размер через ПРОБЕЛ, продолжая строку
    If formattedAmount <> "" And formattedAmount <> "0" Then
        textLine = textLine & " " & t("payments.export.word.amount_prefix", "in amount of") & " " & formattedAmount & " " & t("payments.export.word.position_salary_percent_suffix", "percent of position salary.")
    End If

    ' А вот основание спускаем на новую строку через vbCrLf
    If cleanFound <> "" Then
        textLine = textLine & vbCrLf & FoundationLabel() & ": " & cleanFound
    End If

    FormatEmployeePaymentText = textLine
End Function

