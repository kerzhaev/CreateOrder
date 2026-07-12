Attribute VB_Name = "mdlPaymentPackageSupport"
Option Explicit

Private Const PACKAGE_MODE_LIST As String = "LIST"
Private Const PACKAGE_MODE_INDIVIDUAL As String = "INDIVIDUAL"
Private Const PREVIEW_SHEET_NAME As String = "Payment_Preview"

Public Sub EnsurePaymentsFeatureInfrastructure()
    Dim wsPayments As Worksheet
    Dim wsRef As Worksheet

    On Error Resume Next
    Set wsPayments = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS)
    Set wsRef = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_REF_PAYMENT_TYPES)
    On Error GoTo 0

    If Not wsPayments Is Nothing Then
        EnsurePaymentsSheetEnhancements wsPayments
        EnsurePaymentsSheetButtons wsPayments
    End If

    If Not wsRef Is Nothing Then
        EnsurePaymentTypeReferenceEnhancements wsRef
    End If
End Sub

Public Sub EnsurePaymentsSheetEnhancements(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim widths As Variant
    Dim i As Long

    headers = Array( _
        PT("payments.header.package_id", "Package ID"), _
        PT("payments.header.mode", "Mode"), _
        PT("payments.header.parameter", "Parameter"), _
        PT("payments.header.shared_basis", "Shared basis"), _
        PT("payments.header.group_export", "Grouped export"), _
        PT("payments.header.note", "Note"), _
        PT("payments.header.status", "Status"), _
        PT("payments.header.source_enrollment_id", "Enrollment ID"))
    widths = Array(16, 12, 18, 34, 14, 24, 24, 14)

    For i = LBound(headers) To UBound(headers)
        If Trim$(CStr(ws.Cells(1, mdlPaymentValidation.COL_PACKAGE_ID + i).Value)) = "" Then
            ws.Cells(1, mdlPaymentValidation.COL_PACKAGE_ID + i).Value = headers(i)
        End If
        ws.Columns(mdlPaymentValidation.COL_PACKAGE_ID + i).ColumnWidth = widths(i)
    Next i
End Sub

Public Sub EnsurePaymentsSheetButtons(ByVal ws As Worksheet)
    RemoveSheetButtonIfExists ws, "btnCreatePaymentPackage"
    RemoveSheetButtonIfExists ws, "btnSelectPaymentEmployee"
    RemoveSheetButtonIfExists ws, "btnPastePaymentNumbers"
    RemoveSheetButtonIfExists ws, "btnFillSharedPaymentFields"
    RemoveSheetButtonIfExists ws, "btnRecalcPaymentRows"
    RemoveSheetButtonIfExists ws, "btnPreviewPaymentPackage"
    RemoveSheetButtonIfExists ws, "btnExportPaymentsDocx"
    RemoveSheetButtonIfExists ws, "btnOpenWorkbookFolder"
    RemoveSheetButtonIfExists ws, "btnOpenPaymentsRibbonHint"
End Sub

Private Sub RemoveSheetButtonIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If Not shp Is Nothing Then shp.Delete
End Sub

Public Sub OpenWorkbookFolder()
    ThisWorkbook.FollowHyperlink ThisWorkbook.Path
End Sub

Public Sub EnsurePaymentTypeReferenceEnhancements(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim foundClassQualification As Boolean
    Dim foundFizo As Boolean
    Dim foundSecrecy As Boolean
    Dim i As Long

    If Trim$(CStr(ws.Cells(1, 5).Value)) = "" Then
        ws.Cells(1, 5).Value = "RuleCode"
        ws.Columns(5).ColumnWidth = 20
    End If

    If Trim$(CStr(ws.Cells(1, 13).Value)) = "" Then
        ws.Cells(1, 13).Value = "PaymentEligibilityRule"
        ws.Columns(13).ColumnWidth = 24
    End If

    If Trim$(CStr(ws.Cells(1, 14).Value)) = "" Then
        ws.Cells(1, 14).Value = "PaymentEligibilitySeverity"
        ws.Columns(14).ColumnWidth = 20
    End If

    If Trim$(CStr(ws.Cells(1, 17).Value)) = "" Then
        ws.Cells(1, 17).Value = "PaymentEligibilityPositionKeywords"
        ws.Columns(17).ColumnWidth = 24
    End If

    If Trim$(CStr(ws.Cells(1, 18).Value)) = "" Then
        ws.Cells(1, 18).Value = "PaymentEligibilityFoundationKeywords"
        ws.Columns(18).ColumnWidth = 28
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), PaymentTypeClassQualification(), vbTextCompare) = 0 Then
            foundClassQualification = True
            If Trim$(CStr(ws.Cells(i, 2).Value)) = "" Then ws.Cells(i, 2).Value = "CLASS_QUAL"
            If Trim$(CStr(ws.Cells(i, 3).Value)) = "" Then ws.Cells(i, 3).Value = mdlPaymentTypes.DEFAULT_TEMPLATE
            If Trim$(CStr(ws.Cells(i, 4).Value)) = "" Then ws.Cells(i, 4).Value = PaymentTypeClassQualificationDescription()
            ws.Cells(i, 5).Value = "CLASS_QUAL"
            ws.Cells(i, 13).Value = "PARAM_REQUIRED"
            ws.Cells(i, 14).Value = "WARNING"
            Exit For
        End If
    Next i

    For i = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), PaymentTypeFizo(), vbTextCompare) = 0 Then
            foundFizo = True
            If Trim$(CStr(ws.Cells(i, 13).Value)) = "" Then ws.Cells(i, 13).Value = "PARAM_REQUIRED"
            If Trim$(CStr(ws.Cells(i, 14).Value)) = "" Then ws.Cells(i, 14).Value = "WARNING"
            Exit For
        End If
    Next i

    For i = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), PaymentTypeSecrecy(), vbTextCompare) = 0 Then
            foundSecrecy = True
            If Trim$(CStr(ws.Cells(i, 13).Value)) = "" Then ws.Cells(i, 13).Value = "PARAM_REQUIRED"
            If Trim$(CStr(ws.Cells(i, 14).Value)) = "" Then ws.Cells(i, 14).Value = "BLOCKED"
            If Trim$(CStr(ws.Cells(i, 18).Value)) = "" Then ws.Cells(i, 18).Value = "форма;допуск;секрет"
            Exit For
        End If
    Next i

    If Not foundClassQualification Then
        lastRow = lastRow + 1
        ws.Cells(lastRow, 1).Value = PaymentTypeClassQualification()
        ws.Cells(lastRow, 2).Value = "CLASS_QUAL"
        ws.Cells(lastRow, 3).Value = mdlPaymentTypes.DEFAULT_TEMPLATE
        ws.Cells(lastRow, 4).Value = PaymentTypeClassQualificationDescription()
        ws.Cells(lastRow, 5).Value = "CLASS_QUAL"
        ws.Cells(lastRow, 13).Value = "PARAM_REQUIRED"
        ws.Cells(lastRow, 14).Value = "WARNING"
    End If
End Sub

Public Sub AssignPackageIdToSelection()
    Dim ws As Worksheet
    Dim selectedRange As Range
    Dim rowArea As Range
    Dim packageId As String
    Dim rowNum As Long

    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox PT("payments.message.go_to_sheet", "Open the payments sheet") & " '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, PT("payments.caption.package", "Package")
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set selectedRange = Selection
    If selectedRange Is Nothing Then Exit Sub

    packageId = "PKG-" & Format(Now, "yyyymmdd-hhnnss")

    For Each rowArea In selectedRange.Rows
        rowNum = rowArea.Row
        If rowNum >= 2 Then
            ws.Cells(rowNum, mdlPaymentValidation.COL_PACKAGE_ID).Value = packageId
            ws.Cells(rowNum, mdlPaymentValidation.COL_PACKAGE_MODE).Value = PACKAGE_MODE_LIST
            ws.Cells(rowNum, mdlPaymentValidation.COL_GROUP_EXPORT).Value = GroupExportEnabledValue()
        End If
    Next rowArea

    MsgBox PT("payments.message.package_created", "Package created: ") & packageId, vbInformation, PT("payments.caption.package", "Package")
    Exit Sub

ErrorHandler:
    MsgBox PT("payments.message.package_create_error", "Error while creating package: ") & Err.Description, vbCritical, PT("payments.caption.package", "Package")
End Sub

Public Sub RecalculateSelectedPaymentRows()
    Dim ws As Worksheet
    Dim selectedRange As Range
    Dim rowArea As Range

    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox PT("payments.message.go_to_sheet", "Open the payments sheet") & " '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, PT("payments.caption.recalc", "Recalculate")
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set selectedRange = Selection
    If selectedRange Is Nothing Then Exit Sub

    For Each rowArea In selectedRange.Rows
        If rowArea.Row >= 2 Then
            EnrichPaymentRow ws, rowArea.Row
        End If
    Next rowArea

    MsgBox PT("payments.message.rows_recalculated", "Selected rows recalculated."), vbInformation, PT("payments.caption.recalc", "Recalculate")
    Exit Sub

ErrorHandler:
    MsgBox PT("payments.message.recalc_error", "Error while recalculating rows: ") & Err.Description, vbCritical, PT("payments.caption.recalc", "Recalculate")
End Sub

Public Sub BulkFillSelectedPaymentRows()
    Dim ws As Worksheet
    Dim selectedRows As Collection
    Dim firstRow As Long
    Dim packageId As String
    Dim packageMode As String
    Dim paymentType As String
    Dim parameterValue As String
    Dim sharedBasis As String
    Dim groupExportFlag As String
    Dim wasCancelled As Boolean

    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox PT("payments.message.go_to_sheet", "Open the payments sheet") & " '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, PT("payments.caption.fill", "Fill shared fields")
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set selectedRows = GetSelectedPaymentRows()
    If selectedRows.Count = 0 Then
        MsgBox PT("payments.message.no_rows_selected", "Select payment rows first."), vbExclamation, PT("payments.caption.fill", "Fill shared fields")
        Exit Sub
    End If

    firstRow = CLng(selectedRows(1))

    packageId = PromptTextValue( _
        PT("payments.prompt.package_id", "Package ID (leave current value or enter a new one):"), _
        PT("payments.caption.fill", "Fill shared fields"), _
        Trim$(CStr(ws.Cells(firstRow, mdlPaymentValidation.COL_PACKAGE_ID).Value)), _
        wasCancelled)
    If wasCancelled Then Exit Sub

    packageMode = PromptTextValue( _
        PT("payments.prompt.package_mode", "Package mode: LIST or INDIVIDUAL"), _
        PT("payments.caption.fill", "Fill shared fields"), _
        Trim$(CStr(ws.Cells(firstRow, mdlPaymentValidation.COL_PACKAGE_MODE).Value)), _
        wasCancelled)
    If wasCancelled Then Exit Sub

    paymentType = PromptTextValue( _
        PT("payments.prompt.payment_type", "Payment type for selected rows:"), _
        PT("payments.caption.fill", "Fill shared fields"), _
        Trim$(CStr(ws.Cells(firstRow, mdlPaymentValidation.COL_PAYMENT_TYPE).Value)), _
        wasCancelled)
    If wasCancelled Then Exit Sub

    parameterValue = PromptTextValue( _
        PT("payments.prompt.parameter", "Common parameter for selected rows (optional):"), _
        PT("payments.caption.fill", "Fill shared fields"), _
        Trim$(CStr(ws.Cells(firstRow, mdlPaymentValidation.COL_PARAMETER).Value)), _
        wasCancelled)
    If wasCancelled Then Exit Sub

    sharedBasis = PromptTextValue( _
        PT("payments.prompt.shared_basis", "Shared basis text for selected rows:"), _
        PT("payments.caption.fill", "Fill shared fields"), _
        Trim$(CStr(ws.Cells(firstRow, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value)), _
        wasCancelled)
    If wasCancelled Then Exit Sub

    groupExportFlag = PromptTextValue( _
        PT("payments.prompt.group_export", "Grouped export flag: YES/NO"), _
        PT("payments.caption.fill", "Fill shared fields"), _
        Trim$(CStr(ws.Cells(firstRow, mdlPaymentValidation.COL_GROUP_EXPORT).Value)), _
        wasCancelled)
    If wasCancelled Then Exit Sub

    packageMode = NormalizePackageModeInput(packageMode)
    groupExportFlag = NormalizeGroupExportValue(groupExportFlag)

    If selectedRows.Count > 1 Then
        If packageId = "" And (packageMode = PACKAGE_MODE_LIST Or groupExportFlag = GroupExportEnabledValue()) Then
            packageId = BuildGeneratedPackageId()
        End If
        If packageMode = "" Then packageMode = PACKAGE_MODE_LIST
        If groupExportFlag = "" Then groupExportFlag = GroupExportEnabledValue()
    End If

    ApplySharedValuesToRows ws, selectedRows, packageId, packageMode, paymentType, parameterValue, sharedBasis, groupExportFlag

    MsgBox PT("payments.message.fill_done", "Shared values were applied to selected rows."), vbInformation, PT("payments.caption.fill", "Fill shared fields")
    Exit Sub

ErrorHandler:
    MsgBox PT("payments.message.fill_error", "Error while filling shared values: ") & Err.Description, vbCritical, PT("payments.caption.fill", "Fill shared fields")
End Sub

Public Sub ApplySharedValuesToRange(ByVal sheetName As String, ByVal rangeAddress As String, ByVal packageId As String, ByVal packageMode As String, ByVal paymentType As String, ByVal parameterValue As String, ByVal sharedBasis As String, ByVal groupExportFlag As String)
    Dim ws As Worksheet
    Dim selectedRows As Collection

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set selectedRows = GetPaymentRowsFromRange(ws.Range(rangeAddress))
    ApplySharedValuesToRows ws, selectedRows, packageId, NormalizePackageModeInput(packageMode), paymentType, parameterValue, sharedBasis, NormalizeGroupExportValue(groupExportFlag)
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "ApplySharedValuesToRange", Err.Description
End Sub

Public Sub PreviewSelectedPaymentRows()
    Dim ws As Worksheet
    Dim selectedRows As Collection
    Dim previewText As String

    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox PT("payments.message.go_to_sheet", "Open the payments sheet") & " '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, PT("payments.caption.preview", "Preview")
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set selectedRows = GetSelectedPaymentRows()
    If selectedRows.Count = 0 Then
        MsgBox PT("payments.message.no_rows_selected", "Select payment rows first."), vbExclamation, PT("payments.caption.preview", "Preview")
        Exit Sub
    End If

    PreparePreviewForRows ws, selectedRows
    MsgBox PT("payments.message.preview_ready", "Preview was prepared on the preview sheet."), vbInformation, PT("payments.caption.preview", "Preview")
    Exit Sub

ErrorHandler:
    MsgBox PT("payments.message.preview_error", "Error while preparing preview: ") & Err.Description, vbCritical, PT("payments.caption.preview", "Preview")
End Sub

Public Sub SelectEmployeeForActivePaymentRow()
    Dim wsPayments As Worksheet
    Dim activeCell As Range
    Dim targetRow As Long

    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox PT("payments.message.go_to_sheet", "Open the payments sheet") & " '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, PT("payments.caption.select_employee", "Select employee")
        Exit Sub
    End If

    Set wsPayments = ActiveSheet
    On Error Resume Next
    Set activeCell = Application.ActiveCell
    On Error GoTo ErrorHandler

    If activeCell Is Nothing Then Exit Sub
    If activeCell.Row < 2 Then
        MsgBox PT("payments.message.choose_data_row", "Select a data row on the payments sheet."), vbExclamation, PT("payments.caption.select_employee", "Select employee")
        Exit Sub
    End If

    targetRow = activeCell.Row
    frmSelectEmployee.selectedLichniyNomer = ""
    frmSelectEmployee.selectedFIO = ""
    frmSelectEmployee.isCancelled = True
    frmSelectEmployee.Show

    If Not frmSelectEmployee.isCancelled Then
        FillPaymentEmployeeRow wsPayments, targetRow, frmSelectEmployee.selectedLichniyNomer, frmSelectEmployee.selectedFIO
    End If
    Exit Sub

ErrorHandler:
    MsgBox PT("payments.message.select_employee_error", "Error while selecting employee: ") & Err.Description, vbCritical, PT("payments.caption.select_employee", "Select employee")
End Sub

Public Sub ImportEmployeesByNumberList()
    Dim ws As Worksheet
    Dim activeCell As Range
    Dim numbersText As String
    Dim startRow As Long
    Dim importedCount As Long
    Dim wasCancelled As Boolean

    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox PT("payments.message.go_to_sheet", "Open the payments sheet") & " '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, PT("payments.caption.import_numbers", "Paste number list")
        Exit Sub
    End If

    Set ws = ActiveSheet
    On Error Resume Next
    Set activeCell = Application.ActiveCell
    On Error GoTo ErrorHandler
    If activeCell Is Nothing Then Exit Sub

    startRow = activeCell.Row
    If startRow < 2 Then startRow = 2

    numbersText = PromptTextValue( _
        PT("payments.prompt.number_list", "Paste personal or table numbers. Use a new line, comma, semicolon, or space as a separator:"), _
        PT("payments.caption.import_numbers", "Paste number list"), _
        "", _
        wasCancelled)
    If wasCancelled Then Exit Sub

    importedCount = ImportEmployeesFromTextToSheet(ws.Name, startRow, numbersText)
    MsgBox tf("payments.message.import_numbers_done", "Imported employees: {count}", "{count}", importedCount), vbInformation, PT("payments.caption.import_numbers", "Paste number list")
    Exit Sub

ErrorHandler:
    MsgBox PT("payments.message.import_numbers_error", "Error while importing number list: ") & Err.Description, vbCritical, PT("payments.caption.import_numbers", "Paste number list")
End Sub

Public Sub ImportEmployeesFromStaffSelection()
    Dim wsStaff As Worksheet
    Dim wsPayments As Worksheet
    Dim selectedRange As Range
    Dim startRowText As String
    Dim startRow As Long
    Dim importedCount As Long
    Dim wasCancelled As Boolean

    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_STAFF Then
        MsgBox PT("payments.message.go_to_staff_sheet", "Open the staff sheet and select employee rows first."), vbExclamation, PT("payments.caption.import_from_staff", "Copy from staff")
        Exit Sub
    End If

    Set wsStaff = ActiveSheet
    Set wsPayments = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS)
    On Error Resume Next
    Set selectedRange = Selection
    On Error GoTo ErrorHandler
    If selectedRange Is Nothing Then Exit Sub

    startRowText = PromptTextValue( _
        PT("payments.prompt.import_start_row", "Target row on the payments sheet:"), _
        PT("payments.caption.import_from_staff", "Copy from staff"), _
        CStr(GetSuggestedPaymentInsertRow(wsPayments)), _
        wasCancelled)
    If wasCancelled Then Exit Sub

    If Not IsNumeric(startRowText) Then
        Err.Raise vbObjectError + 741, "ImportEmployeesFromStaffSelection", PT("payments.message.invalid_start_row", "Enter a numeric target row.")
    End If

    startRow = CLng(startRowText)
    importedCount = ImportEmployeesFromStaffRange(wsStaff.Name, wsPayments.Name, selectedRange.Address(False, False), startRow)

    wsPayments.Activate
    wsPayments.Cells(startRow, mdlPaymentValidation.COL_FIO).Select

    MsgBox tf("payments.message.import_from_staff_done", "Copied employees from staff: {count}", "{count}", importedCount), vbInformation, PT("payments.caption.import_from_staff", "Copy from staff")
    Exit Sub

ErrorHandler:
    MsgBox PT("payments.message.import_from_staff_error", "Error while copying employees from staff: ") & Err.Description, vbCritical, PT("payments.caption.import_from_staff", "Copy from staff")
End Sub

Public Function ImportEmployeesFromStaffRange(ByVal staffSheetName As String, ByVal paymentsSheetName As String, ByVal rangeAddress As String, ByVal startRow As Long) As Long
    Dim wsStaff As Worksheet
    Dim wsPayments As Worksheet
    Dim selectedRows As Collection
    Dim i As Long
    Dim sourceRow As Long
    Dim targetRow As Long
    Dim anchorRow As Long

    On Error GoTo ErrorHandler

    mdlHelper.EnsureStaffColumnsInitialized

    Set wsStaff = ThisWorkbook.Worksheets(staffSheetName)
    Set wsPayments = ThisWorkbook.Worksheets(paymentsSheetName)
    Set selectedRows = GetPaymentRowsFromRange(wsStaff.Range(rangeAddress))

    If startRow < 2 Then startRow = 2
    If selectedRows.Count = 0 Then Exit Function

    anchorRow = startRow

    For i = 1 To selectedRows.Count
        sourceRow = CLng(selectedRows(i))
        targetRow = startRow + i - 1

        If i > 1 Then
            CopySharedPaymentFields wsPayments, anchorRow, targetRow
        End If

        FillPaymentEmployeeRow wsPayments, targetRow, _
            CStr(wsStaff.Cells(sourceRow, mdlHelper.colLichniyNomer_Global).Value), _
            CStr(wsStaff.Cells(sourceRow, mdlHelper.colFIO_Global).Value)
    Next i

    FillPaymentSequenceNumbers wsPayments
    ImportEmployeesFromStaffRange = selectedRows.Count
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ImportEmployeesFromStaffRange", Err.Description
End Function

Public Function ImportEmployeesFromTextToSheet(ByVal sheetName As String, ByVal startRow As Long, ByVal numbersText As String) As Long
    Dim ws As Worksheet
    Dim tokens As Collection
    Dim i As Long
    Dim rowNum As Long
    Dim staffData As Object
    Dim anchorRow As Long

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(sheetName)
    If startRow < 2 Then startRow = 2

    Set tokens = ParseEmployeeNumberTokens(numbersText)
    If tokens.Count = 0 Then Exit Function

    anchorRow = startRow

    For i = 1 To tokens.Count
        rowNum = startRow + i - 1
        If i > 1 Then
            CopySharedPaymentFields ws, anchorRow, rowNum
        End If

        Set staffData = mdlHelper.FindEmployeeByAnyNumber(CStr(tokens(i)))
        If staffData Is Nothing Then
            FillPaymentEmployeeRow ws, rowNum, CStr(tokens(i)), ""
        ElseIf staffData.Count > 0 Then
            FillPaymentEmployeeRow ws, rowNum, CStr(staffData("Личный номер")), CStr(staffData("Лицо"))
        Else
            FillPaymentEmployeeRow ws, rowNum, CStr(tokens(i)), ""
        End If
    Next i

    FillPaymentSequenceNumbers ws
    ImportEmployeesFromTextToSheet = tokens.Count
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ImportEmployeesFromTextToSheet", Err.Description
End Function

Public Sub PreparePreviewForRange(ByVal sheetName As String, ByVal rangeAddress As String)
    Dim ws As Worksheet
    Dim selectedRows As Collection

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set selectedRows = GetPaymentRowsFromRange(ws.Range(rangeAddress))
    PreparePreviewForRows ws, selectedRows
End Sub

Private Sub ApplySharedValuesToRows(ByVal ws As Worksheet, ByVal selectedRows As Collection, ByVal packageId As String, ByVal packageMode As String, ByVal paymentType As String, ByVal parameterValue As String, ByVal sharedBasis As String, ByVal groupExportFlag As String)
    Dim i As Long
    Dim rowNum As Long
    Dim currentFoundation As String
    Dim currentSharedBasis As String

    For i = 1 To selectedRows.Count
        rowNum = CLng(selectedRows(i))
        If rowNum >= 2 Then
            currentFoundation = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_FOUNDATION).Value))
            currentSharedBasis = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value))

            ws.Cells(rowNum, mdlPaymentValidation.COL_PACKAGE_ID).Value = packageId
            ws.Cells(rowNum, mdlPaymentValidation.COL_PACKAGE_MODE).Value = packageMode
            ws.Cells(rowNum, mdlPaymentValidation.COL_PAYMENT_TYPE).Value = paymentType
            ws.Cells(rowNum, mdlPaymentValidation.COL_PARAMETER).Value = parameterValue
            ws.Cells(rowNum, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value = sharedBasis
            ws.Cells(rowNum, mdlPaymentValidation.COL_GROUP_EXPORT).Value = groupExportFlag

            If currentFoundation = "" Or currentFoundation = currentSharedBasis Then
                ws.Cells(rowNum, mdlPaymentValidation.COL_FOUNDATION).Value = sharedBasis
            End If

            EnrichPaymentRow ws, rowNum
        End If
    Next i
End Sub

Private Sub CopySharedPaymentFields(ByVal ws As Worksheet, ByVal sourceRow As Long, ByVal targetRow As Long)
    If sourceRow < 2 Or targetRow < 2 Then Exit Sub

    ws.Cells(targetRow, mdlPaymentValidation.COL_PAYMENT_TYPE).Value = ws.Cells(sourceRow, mdlPaymentValidation.COL_PAYMENT_TYPE).Value
    ws.Cells(targetRow, mdlPaymentValidation.COL_FOUNDATION).Value = ws.Cells(sourceRow, mdlPaymentValidation.COL_FOUNDATION).Value
    ws.Cells(targetRow, mdlPaymentValidation.COL_PACKAGE_ID).Value = ws.Cells(sourceRow, mdlPaymentValidation.COL_PACKAGE_ID).Value
    ws.Cells(targetRow, mdlPaymentValidation.COL_PACKAGE_MODE).Value = ws.Cells(sourceRow, mdlPaymentValidation.COL_PACKAGE_MODE).Value
    ws.Cells(targetRow, mdlPaymentValidation.COL_PARAMETER).Value = ws.Cells(sourceRow, mdlPaymentValidation.COL_PARAMETER).Value
    ws.Cells(targetRow, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value = ws.Cells(sourceRow, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value
    ws.Cells(targetRow, mdlPaymentValidation.COL_GROUP_EXPORT).Value = ws.Cells(sourceRow, mdlPaymentValidation.COL_GROUP_EXPORT).Value
End Sub

Private Sub FillPaymentEmployeeRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal personalNumber As String, ByVal fioText As String)
    ws.Cells(rowNum, mdlPaymentValidation.COL_LICHNIY_NOMER).Value = Trim$(personalNumber)
    ws.Cells(rowNum, mdlPaymentValidation.COL_FIO).Value = Trim$(fioText)
    EnrichPaymentRow ws, rowNum
End Sub

Private Sub FillPaymentSequenceNumbers(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rowNum As Long

    lastRow = ws.Cells(ws.Rows.Count, mdlPaymentValidation.COL_LICHNIY_NOMER).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For rowNum = 2 To lastRow
        If Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_LICHNIY_NOMER).Value)) <> "" Then
            ws.Cells(rowNum, mdlPaymentValidation.COL_NUMBER).Value = rowNum - 1
        End If
    Next rowNum
End Sub

Private Function ParseEmployeeNumberTokens(ByVal numbersText As String) As Collection
    Dim result As Collection
    Dim normalizedText As String
    Dim rawTokens() As String
    Dim token As Variant
    Dim cleanToken As String

    Set result = New Collection
    normalizedText = Replace$(numbersText, vbCr, vbLf)
    normalizedText = Replace$(normalizedText, vbTab, vbLf)
    normalizedText = Replace$(normalizedText, ";", vbLf)
    normalizedText = Replace$(normalizedText, ",", vbLf)
    normalizedText = Replace$(normalizedText, " ", vbLf)

    Do While InStr(normalizedText, vbLf & vbLf) > 0
        normalizedText = Replace$(normalizedText, vbLf & vbLf, vbLf)
    Loop

    rawTokens = Split(normalizedText, vbLf)
    For Each token In rawTokens
        cleanToken = Trim$(CStr(token))
        If cleanToken <> "" Then
            result.Add cleanToken
        End If
    Next token

    Set ParseEmployeeNumberTokens = result
End Function

Private Function GetSelectedPaymentRows() As Collection
    On Error Resume Next
    Set GetSelectedPaymentRows = GetPaymentRowsFromRange(Selection)
    On Error GoTo 0
    If GetSelectedPaymentRows Is Nothing Then Set GetSelectedPaymentRows = New Collection
End Function

Private Function GetSuggestedPaymentInsertRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, mdlPaymentValidation.COL_LICHNIY_NOMER).End(xlUp).Row
    If lastRow < 2 Then
        GetSuggestedPaymentInsertRow = 2
    ElseIf Trim$(CStr(ws.Cells(lastRow, mdlPaymentValidation.COL_LICHNIY_NOMER).Value)) = "" Then
        GetSuggestedPaymentInsertRow = lastRow
    Else
        GetSuggestedPaymentInsertRow = lastRow + 1
    End If
End Function

Private Function GetPaymentRowsFromRange(ByVal selectedRange As Range) As Collection
    Dim result As Collection
    Dim seenRows As Object
    Dim rowArea As Range
    Dim rowKey As String

    Set result = New Collection
    If selectedRange Is Nothing Then
        Set GetPaymentRowsFromRange = result
        Exit Function
    End If

    Set seenRows = CreateObject("Scripting.Dictionary")

    For Each rowArea In selectedRange.Rows
        If rowArea.Row >= 2 Then
            rowKey = CStr(rowArea.Row)
            If Not seenRows.Exists(rowKey) Then
                seenRows.Add rowKey, True
                result.Add CLng(rowArea.Row)
            End If
        End If
    Next rowArea

    Set GetPaymentRowsFromRange = result
End Function

Private Function PromptTextValue(ByVal promptText As String, ByVal titleText As String, ByVal defaultValue As String, ByRef wasCancelled As Boolean) As String
    Dim response As Variant

    response = Application.InputBox(promptText, titleText, defaultValue, Type:=2)
    If VarType(response) = vbBoolean And response = False Then
        wasCancelled = True
        Exit Function
    End If

    PromptTextValue = CStr(response)
End Function

Private Function NormalizePackageModeInput(ByVal rawValue As String) As String
    Dim normalizedValue As String

    normalizedValue = NormalizeTextValue(rawValue)
    Select Case normalizedValue
        Case "LIST", "СПИСОЧНО", "СПИСОК"
            NormalizePackageModeInput = PACKAGE_MODE_LIST
        Case "INDIVIDUAL", "ИНДИВИДУАЛЬНО", "ОДИНОЧНО"
            NormalizePackageModeInput = PACKAGE_MODE_INDIVIDUAL
        Case Else
            NormalizePackageModeInput = normalizedValue
    End Select
End Function

Private Function NormalizeGroupExportValue(ByVal rawValue As String) As String
    Dim normalizedValue As String

    normalizedValue = NormalizeTextValue(rawValue)
    Select Case normalizedValue
        Case "ДА", "YES", "TRUE", "1"
            NormalizeGroupExportValue = GroupExportEnabledValue()
        Case "НЕТ", "NO", "FALSE", "0"
            NormalizeGroupExportValue = PT("payments.value.group_export_no", "NO")
        Case Else
            NormalizeGroupExportValue = normalizedValue
    End Select
End Function

Private Function BuildGeneratedPackageId() As String
    BuildGeneratedPackageId = "PKG-" & Format(Now, "yyyymmdd-hhnnss")
End Function

Private Sub PreparePreviewForRows(ByVal ws As Worksheet, ByVal selectedRows As Collection)
    Dim previewText As String

    previewText = mdlUniversalPaymentExport.BuildPaymentPreviewTextFromRows(ws, selectedRows, PT("payments.preview.title", "Preview of export text"))
    If Trim$(previewText) = "" Then
        Err.Raise vbObjectError + 721, "PreparePreviewForRows", PT("payments.message.preview_empty", "No data was found for preview.")
    End If

    ShowPreviewOnWorksheet previewText
End Sub

Private Sub ShowPreviewOnWorksheet(ByVal previewText As String)
    Dim wsPreview As Worksheet

    Set wsPreview = Nothing
    On Error Resume Next
    Set wsPreview = ThisWorkbook.Worksheets(PREVIEW_SHEET_NAME)
    On Error GoTo 0

    If wsPreview Is Nothing Then
        Set wsPreview = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsPreview.Name = PREVIEW_SHEET_NAME
    End If

    wsPreview.Cells.Clear
    wsPreview.Range("A1").Value = previewText
    wsPreview.Range("A1").EntireColumn.ColumnWidth = 100
    wsPreview.Range("A1").EntireRow.RowHeight = 20
    wsPreview.Range("A1").WrapText = True
    wsPreview.Range("A1").HorizontalAlignment = xlLeft
    wsPreview.Range("A1").VerticalAlignment = xlTop
    wsPreview.Cells(1, 1).EntireRow.AutoFit
    wsPreview.Activate
    wsPreview.Range("A1").Select
End Sub

Public Sub EnrichPaymentRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    Dim explicitAmount As String
    Dim parameterValue As String
    Dim paymentType As String
    Dim rowFoundation As String
    Dim sharedFoundation As String
    Dim calculatedAmount As String

    paymentType = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_PAYMENT_TYPE).Value))
    explicitAmount = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_AMOUNT).Value))
    parameterValue = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_PARAMETER).Value))
    rowFoundation = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_FOUNDATION).Value))
    sharedFoundation = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value))

    If rowFoundation = "" And sharedFoundation <> "" Then
        ws.Cells(rowNum, mdlPaymentValidation.COL_FOUNDATION).Value = sharedFoundation
    End If

    calculatedAmount = ResolvePaymentAmount(paymentType, parameterValue, explicitAmount)
    If explicitAmount = "" And calculatedAmount <> "" Then
        ws.Cells(rowNum, mdlPaymentValidation.COL_AMOUNT).Value = calculatedAmount
    End If
End Sub

Public Function ResolvePaymentAmount(ByVal paymentType As String, ByVal parameterValue As String, ByVal explicitAmount As String) As String
    Dim config As mdlPaymentTypes.PaymentTypeConfig
    Dim normalizedParameter As String

    If Trim$(explicitAmount) <> "" Then
        ResolvePaymentAmount = Trim$(explicitAmount)
        Exit Function
    End If

    normalizedParameter = NormalizeTextValue(parameterValue)
    If normalizedParameter = "" Then Exit Function

    config = mdlPaymentTypes.GetPaymentTypeConfig(paymentType)

    If StrComp(config.RuleCode, "CLASS_QUAL", vbTextCompare) = 0 Then
        ResolvePaymentAmount = ResolveClassQualificationAmount(normalizedParameter)
        Exit Function
    End If

    If IsSimplePercentText(normalizedParameter) Then
        ResolvePaymentAmount = NormalizePercentText(normalizedParameter)
    End If
End Function

Private Function ResolveClassQualificationAmount(ByVal normalizedParameter As String) As String
    If InStr(1, normalizedParameter, mdlHelper.Ru(1084, 1072, 1089, 1090, 1077, 1088), vbTextCompare) > 0 Then
        ResolveClassQualificationAmount = "30%"
    ElseIf normalizedParameter = "1" Or InStr(1, normalizedParameter, mdlHelper.Ru(49, 32, 1082, 1083, 1072, 1089, 1089), vbTextCompare) > 0 Then
        ResolveClassQualificationAmount = "20%"
    ElseIf normalizedParameter = "2" Or InStr(1, normalizedParameter, mdlHelper.Ru(50, 32, 1082, 1083, 1072, 1089, 1089), vbTextCompare) > 0 Then
        ResolveClassQualificationAmount = "10%"
    ElseIf normalizedParameter = "3" Or InStr(1, normalizedParameter, mdlHelper.Ru(51, 32, 1082, 1083, 1072, 1089, 1089), vbTextCompare) > 0 Then
        ResolveClassQualificationAmount = "5%"
    End If
End Function

Private Function IsSimplePercentText(ByVal valueText As String) As Boolean
    Dim normalizedText As String

    normalizedText = Replace$(Replace$(Trim$(valueText), "%", ""), ",", ".")
    If normalizedText = "" Then Exit Function

    If IsNumeric(normalizedText) Then
        IsSimplePercentText = True
    End If
End Function

Private Function NormalizePercentText(ByVal valueText As String) As String
    Dim normalizedText As String
    Dim numericValue As Double

    normalizedText = Replace$(Replace$(Trim$(valueText), "%", ""), ",", ".")
    If Not IsNumeric(normalizedText) Then Exit Function

    numericValue = CDbl(normalizedText)
    If numericValue > 0 And numericValue <= 1 Then
        numericValue = numericValue * 100
    End If

    NormalizePercentText = Replace$(CStr(numericValue), ",", ".") & "%"
End Function

Public Function GetEffectiveFoundationFromSheet(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim rowFoundation As String
    Dim sharedFoundation As String

    rowFoundation = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_FOUNDATION).Value))
    If rowFoundation <> "" Then
        GetEffectiveFoundationFromSheet = rowFoundation
        Exit Function
    End If

    sharedFoundation = Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value))
    GetEffectiveFoundationFromSheet = sharedFoundation
End Function

Public Function GetEffectiveAmountFromSheet(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    GetEffectiveAmountFromSheet = ResolvePaymentAmount( _
        Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_PAYMENT_TYPE).Value)), _
        Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_PARAMETER).Value)), _
        Trim$(CStr(ws.Cells(rowNum, mdlPaymentValidation.COL_AMOUNT).Value)))
End Function

Public Function BuildExportGroupKey(ByRef payment As mdlPaymentTypes.PaymentWithoutPeriod) As String
    Dim paymentTypeKey As String

    paymentTypeKey = Trim$(LCase$(payment.paymentType))
    If paymentTypeKey = "" Then paymentTypeKey = "без_типа"

    If ShouldUseGroupedExport(payment) And Trim$(payment.packageId) <> "" Then
        BuildExportGroupKey = paymentTypeKey & "|" & Trim$(payment.packageId)
    Else
        BuildExportGroupKey = paymentTypeKey
    End If
End Function

Public Function ShouldUseGroupedExport(ByRef payment As mdlPaymentTypes.PaymentWithoutPeriod) As Boolean
    Dim normalizedFlag As String
    Dim normalizedMode As String

    normalizedFlag = NormalizeTextValue(payment.groupExportFlag)
    normalizedMode = NormalizeTextValue(payment.packageMode)

    If normalizedFlag = "ДА" Or normalizedFlag = "YES" Or normalizedFlag = "TRUE" Or normalizedFlag = "1" Then
        ShouldUseGroupedExport = True
        Exit Function
    End If

    If normalizedMode = PACKAGE_MODE_LIST Or normalizedMode = "СПИСОЧНО" Then
        ShouldUseGroupedExport = True
    End If
End Function

Public Function NormalizeTextValue(ByVal valueText As String) As String
    NormalizeTextValue = UCase$(Trim$(CStr(valueText)))
End Function

Private Function PT(ByVal key As String, ByVal fallback As String) As String
    PT = t(key, fallback)
End Function

Private Function PaymentTypeClassQualification() As String
    PaymentTypeClassQualification = t("payments.type.class_qualification", mdlHelper.Ru(1050, 1083, 1072, 1089, 1089, 1085, 1072, 1103, 32, 1082, 1074, 1072, 1083, 1080, 1092, 1080, 1082, 1072, 1094, 1080, 1103))
End Function

Private Function PaymentTypeFizo() As String
    PaymentTypeFizo = t("payments.type.fizo", mdlHelper.Ru(1060, 1048, 1047, 1054))
End Function

Private Function PaymentTypeSecrecy() As String
    PaymentTypeSecrecy = t("payments.type.secrecy", mdlHelper.Ru(1057, 1077, 1082, 1088, 1077, 1090, 1085, 1086, 1089, 1090, 1100))
End Function

Private Function PaymentTypeClassQualificationDescription() As String
    PaymentTypeClassQualificationDescription = t("payments.type.class_qualification_desc", mdlHelper.Ru(1053, 1072, 1076, 1073, 1072, 1074, 1082, 1072, 32, 1079, 1072, 32, 1082, 1083, 1072, 1089, 1089, 1085, 1091, 1102, 32, 1082, 1074, 1072, 1083, 1080, 1092, 1080, 1082, 1072, 1094, 1080, 1102))
End Function

Private Function GroupExportEnabledValue() As String
    GroupExportEnabledValue = t("payments.value.group_export_yes", mdlHelper.Ru(1044, 1040))
End Function
