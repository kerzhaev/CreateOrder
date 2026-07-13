Attribute VB_Name = "mdlRibbonHandlers"
' ===============================================================================
' Module mdlRibbonHandlers for handling custom ribbon events
' Version: 2.3.0 (Updated UI & Licensing integration)
' Date: 23.02.2026
' Author: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' ===============================================================================
Option Explicit

' ===============================================================================
' RIBBON LOCALIZATION CALLBACKS
' ===============================================================================

Public Sub GetRibbonLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRibbonUiTextById(control.Id, "label")
End Sub

Public Sub GetRibbonScreentip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRibbonUiTextById(control.Id, "screentip")
End Sub

Public Sub GetRibbonSupertip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRibbonUiTextById(control.Id, "supertip")
End Sub

Public Function GetRibbonUiTextById(ByVal controlId As String, Optional ByVal textKind As String = "label") As String
    Dim normalizedKind As String
    Dim fallback As String

    normalizedKind = LCase$(Trim$(textKind))
    If normalizedKind = "" Then normalizedKind = "label"

    fallback = ""
    If normalizedKind = "label" Then fallback = Trim$(controlId)

    GetRibbonUiTextById = t("ribbon.ui." & Trim$(controlId) & "." & normalizedKind, fallback)
End Function

' ===============================================================================
' PREMIUM FUNCTIONS (Require active license / free period)
' ===============================================================================

' Handler for main order
Sub RunMainExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Application.ScreenUpdating = False
    Call mdlMainExport.ExportToWordFromStaffByLichniyNomer
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.main_export", "Ошибка при создании основного приказа: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for DSO certificate (spravka)
Sub RunSpravkaExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Application.ScreenUpdating = False
    Call mdlSpravkaExport.ExportToWordSpravkaFromTemplate
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.spravka_export", "Ошибка при создании справки: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for report (raport) with CHOICE
Sub RunRaportExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Dim choice As VbMsgBoxResult
    choice = MsgBox(t("ribbon.raport.choice_text", "Какой рапорт необходимо сформировать?" & vbCrLf & vbCrLf & _
                    "Да - Рапорт на ДСО (Сутки отдыха)" & vbCrLf & _
                    "Нет - Рапорт на РИСК (Денежная выплата)" & vbCrLf & _
                    "Отмена - Выход"), vbYesNoCancel + vbQuestion, t("ribbon.raport.choice_title", "Выбор типа рапорта"))
    
    If choice = vbCancel Then Exit Sub
    
    Application.ScreenUpdating = False
    
    If choice = vbYes Then
        ' Если нажали ДА - формируем обычный рапорт на отгулы
        Call mdlRaportExport.ExportToWordRaportFromTemplateByLichniyNomer("DSO")
    Else
        ' Если нажали НЕТ - формируем рапорт на риск (2%)
        Call mdlRaportExport.ExportToWordRaportFromTemplateByLichniyNomer("Risk")
    End If
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.raport_export", "Ошибка при создании рапорта: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for "OrderForRisk" button
Public Sub OnRiskOrderClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Call mdlRiskExport.ExportRiskAllowanceOrder
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.risk_order", "Ошибка при вызове приказа за риск: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for "Export Allowances" button
Public Sub OnExportAllowancesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Application.ScreenUpdating = False
    Call mdlUniversalPaymentExport.ExportPaymentsWithoutPeriods
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.allowances_export", "Ошибка при экспорте надбавок: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for Excel reports (Alushta / FRP)
Public Sub OnPeriodsReportClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Call mdlFRPExport.ExportPeriodsToExcel_WithChoice
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.periods_report", "Ошибка при создании Excel отчета: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' ===============================================================================
' FREE FUNCTIONS (No license required)
' ===============================================================================

' Умная валидация: сама понимает, какой лист проверять
Sub RunSmartValidation(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If ActiveSheet Is Nothing Then Exit Sub
    Dim sheetName As String
    sheetName = ActiveSheet.Name

    If RunSmartValidationBySheetName(sheetName, False) = "" Then
        MsgBox tf("ribbon.smart_validation.unsupported_sheet", "Для проверки данных перейдите на лист '{dso}', '{payments}' или '{enrollment}'.", "{dso}", mdlHelper.Ru(1044, 1057, 1054), "{payments}", mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS, "{enrollment}", mdlReferenceData.SHEET_ENROLLMENT), vbInformation, t("ribbon.smart_validation.title", "Умная проверка")
    End If
    Application.StatusBar = False
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox tf("ribbon.error.smart_validation", "Ошибка при проверке данных: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Function RunSmartValidationBySheetName(ByVal sheetName As String, Optional ByVal isSilent As Boolean = False) As String
    If sheetName = mdlHelper.Ru(1044, 1057, 1054) Then
        Application.ScreenUpdating = False
        Application.StatusBar = t("ribbon.status.validate_dso", "Проверка периодов ДСО...")
        Call mdlDataValidation.ValidateMainSheetData(isSilent)
        Application.ScreenUpdating = True
        RunSmartValidationBySheetName = "DSO"
    ElseIf sheetName = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        Application.ScreenUpdating = False
        Call mdlPaymentValidation.ValidatePaymentsWithoutPeriods(isSilent)
        Application.ScreenUpdating = True
        RunSmartValidationBySheetName = "PAYMENTS"
    ElseIf sheetName = mdlReferenceData.SHEET_ENROLLMENT Then
        Application.ScreenUpdating = False
        Application.StatusBar = t("ribbon.status.validate_enrollment", "Проверка листа зачисления...")
        Call mdlEnrollmentWorkflow.ValidateEnrollmentSheet(isSilent)
        Application.ScreenUpdating = True
        RunSmartValidationBySheetName = "ENROLLMENT"
    End If

    Application.StatusBar = False
End Function

' НОВОЕ: Обработчик кнопки "О программе" (Заменил Диагностику)
Sub RunShowAbout(control As IRibbonControl)
    On Error GoTo ErrorHandler
    frmAbout.Show
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.about", "Ошибка при открытии окна программы: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for data import
Sub RunImportData(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call mdlDataImport.ImportDataToStaff
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.data_import", "Ошибка при импорте данных: {error}", "{error}", Err.description), vbCritical, t("ribbon.error.import_title", "Ошибка импорта")
End Sub

' Handler for Word Raport Import
Sub RunWordRaportImport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Application.ScreenUpdating = False
    Application.StatusBar = t("ribbon.status.raport_import_init", "Инициализация импорта рапорта...")
    Call mdlWordImport.ExecuteWordImport
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox tf("ribbon.error.word_import", "Ошибка при вызове импорта: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for "Mass Add Employees" button
Public Sub OnMassImportEmployeesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If ActiveSheet Is Nothing Then Exit Sub

    If ActiveSheet.Name = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        Call mdlPaymentPackageSupport.ImportEmployeesByNumberList
    ElseIf ActiveSheet.Name = mdlReferenceData.SHEET_STAFF Then
        Call mdlPaymentPackageSupport.ImportEmployeesFromStaffSelection
    Else
        MsgBox tf("ribbon.mass_import.unsupported_sheet", "Для массового добавления перейдите на лист '{payments}' или на лист '{staff}'.", "{payments}", mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS, "{staff}", mdlReferenceData.SHEET_STAFF), vbExclamation, t("common.attention", "Внимание")
        Exit Sub
    End If
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.mass_import", "Ошибка при массовом добавлении сотрудников: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for "Select Employee" button
Public Sub OnSelectEmployeeClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Call mdlPaymentPackageSupport.SelectEmployeeForActivePaymentRow
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.select_employee", "Ошибка при выборе сотрудника: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for "References" button
Public Sub OnManageReferencesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Dim wsRef As Worksheet
    On Error Resume Next
    Set wsRef = ThisWorkbook.Sheets(mdlReferenceData.SHEET_REF_PAYMENT_TYPES)
    If wsRef Is Nothing Then
        MsgBox t("ribbon.references.not_found", "Лист справочников не найден."), vbInformation, t("ribbon.references.title", "Справочники")
    Else
        wsRef.Activate
        wsRef.Cells(1, 1).Select
    End If
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.references", "Ошибка при открытии справочников: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnOpenPaymentsSheetClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Cells(2, 1).Select
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.open_payments_sheet", "Ошибка при открытии листа выплат: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnPaymentsSelectEmployeeClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    mdlPaymentPackageSupport.SelectEmployeeForActivePaymentRow
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.select_employee", "Ошибка при выборе сотрудника: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnPaymentsPasteNumbersClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    mdlPaymentPackageSupport.ImportEmployeesByNumberList
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.paste_numbers", "Ошибка при вставке списка номеров: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnPaymentsCreatePackageClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    mdlPaymentPackageSupport.AssignPackageIdToSelection
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.create_package", "Ошибка при создании пакета: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnPaymentsFillSharedClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    mdlPaymentPackageSupport.BulkFillSelectedPaymentRows
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.fill_shared", "Ошибка при заполнении общих полей: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnPaymentsRecalcClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    mdlPaymentPackageSupport.RecalculateSelectedPaymentRows
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.recalc_payments", "Ошибка при пересчете строк: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnPaymentsPreviewClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    mdlPaymentPackageSupport.PreviewSelectedPaymentRows
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.payments_preview", "Ошибка при построении предпросмотра: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnPaymentsExportDocxClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS).Activate
    mdlUniversalPaymentExport.ExportPaymentsWithoutPeriods
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.payments_export", "Ошибка при экспорте выплат: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnOpenWorkbookFolderClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    mdlPaymentPackageSupport.OpenWorkbookFolder
    Exit Sub
ErrorHandler:
    MsgBox tf("ribbon.error.open_workbook_folder", "Ошибка при открытии папки книги: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnOpenEnrollmentFormClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    mdlEnrollmentWorkflow.OpenEnrollmentForm
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.open_form", "Ошибка открытия мастера зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnOpenPersonnelEnrollmentActionClick(control As IRibbonControl)
    mdlPersonnelEvents.OpenPersonnelEnrollmentAction
End Sub

Public Sub OnOpenPersonnelTransferActionClick(control As IRibbonControl)
    mdlPersonnelEvents.OpenPersonnelTransferAction
End Sub

Public Sub OnOpenPersonnelExclusionActionClick(control As IRibbonControl)
    mdlPersonnelEvents.OpenPersonnelExclusionAction
End Sub

Public Sub OnOpenPersonnelHistoryActionClick(control As IRibbonControl)
    mdlPersonnelEvents.OpenHistoryForPersonnelAction
End Sub

Public Sub OnSavePersonnelActionClick(control As IRibbonControl)
    mdlPersonnelEvents.SaveCurrentPersonnelAction
End Sub
Public Sub OnExportSavedPersonnelActionClick(control As IRibbonControl)
    mdlPersonnelEvents.ExportSavedPersonnelEventOrder
End Sub
Public Sub OnEnrollmentRefreshFormClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    mdlEnrollmentWorkflow.RefreshEnrollmentForm
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.preview", "Ошибка проверки карточки зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentSaveFormClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim targetRow As Long
    Dim orderDraftId As String
    targetRow = mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet(False)
    orderDraftId = mdlEnrollmentOrderExport.GetOrderDraftIdForRow(targetRow)
    MsgBox tf("enrollment.ribbon.saved", "Карточка зачисления сохранена в строку {row}.{nl}OrderDraftId: {draftId}", "{row}", targetRow, "{nl}", vbCrLf, "{draftId}", orderDraftId), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.save", "Ошибка сохранения карточки зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentSaveAndGeneratePaymentsClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim targetRow As Long
    Dim createdCount As Long
    Dim orderDraftId As String

    targetRow = mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet(False)
    createdCount = mdlEnrollmentWorkflow.GeneratePaymentsFromEnrollmentRowDirect(targetRow)
    orderDraftId = mdlEnrollmentOrderExport.GetOrderDraftIdForRow(targetRow)
    MsgBox tf("enrollment.form.message.saved_generated", "Карточка сохранена в строку {row}. OrderDraftId: {draftId}.{nl}Подготовлено выплат: {count}.", "{row}", targetRow, "{draftId}", orderDraftId, "{nl}", vbCrLf, "{count}", createdCount), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.form.error.save_generate", "Ошибка сохранения карточки зачисления с подготовкой выплат: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentSaveAndPrepareClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim targetRow As Long
    Dim orderDraftId As String
    Dim exportScope As String
    Dim outputPath As String
    Dim exportErrorText As String
    targetRow = mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet(False)
    orderDraftId = mdlEnrollmentOrderExport.GetOrderDraftIdForRow(targetRow)
    exportScope = mdlEnrollmentOrderExport.GetExportScopeText(orderDraftId, targetRow)
    exportErrorText = mdlEnrollmentOrderExport.GetEnrollmentExportBlockingIssues(orderDraftId, targetRow)
    If exportErrorText <> "" Then
        MsgBox tf("enrollment.form.message.export_blocked", "Word-приказ не сформирован.{nl}{error}", "{nl}", vbCrLf, "{error}", exportErrorText), vbExclamation, t("enrollment.caption.main", "Зачисление")
        Exit Sub
    End If
    outputPath = mdlEnrollmentOrderExport.ExportEnrollmentOrderByDraftId(orderDraftId, targetRow)
    If mdlEnrollmentOrderExport.IsEnrollmentExportErrorResult(outputPath) Then
        MsgBox tf("enrollment.form.message.export_blocked", "Word-приказ не сформирован.{nl}{error}", "{nl}", vbCrLf, "{error}", mdlEnrollmentOrderExport.GetEnrollmentExportErrorText(outputPath)), vbExclamation, t("enrollment.caption.main", "Зачисление")
        Exit Sub
    End If
    MsgBox tf("enrollment.ribbon.exported", "Пакет приказа сформирован ({scope}).{nl}Файл: {path}", "{scope}", exportScope, "{nl}", vbCrLf, "{path}", outputPath), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.export", "Ошибка экспорта приказа о зачислении: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentClearFormClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    mdlEnrollmentWorkflow.ClearEnrollmentForm
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.clear", "Ошибка очистки мастера зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnOpenEnrollmentSheetClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT).Activate
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT).Cells(2, 1).Select
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.open_sheet", "Ошибка открытия журнала зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentOpenSelectedClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim rowNum As Long
    rowNum = mdlEnrollmentWorkflow.LoadSelectedEnrollmentRowToBackend()
    frmEnrollmentWizard.Show
    MsgBox tf("enrollment.ribbon.selected_loaded", "Строка {row} загружена в мастер зачисления.", "{row}", rowNum), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.open_selected", "Ошибка открытия выбранной строки зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentSaveAndContinuePackageClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim orderDraftId As String

    orderDraftId = mdlEnrollmentWorkflow.SaveEnrollmentFormAndContinuePackage()
    frmEnrollmentWizard.Show
    MsgBox tf("enrollment.form.message.package_next", "Следующий военнослужащий пакета подготовлен. OrderDraftId: {draftId}", "{draftId}", orderDraftId), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.package_next", "Ошибка подготовки следующей карточки пакета: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentContinuePackageClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim rowNum As Long
    Dim orderDraftId As String

    rowNum = mdlEnrollmentWorkflow.ResolveActiveEnrollmentRow()
    orderDraftId = mdlEnrollmentWorkflow.PrepareNextEnrollmentInPackage(rowNum)
    frmEnrollmentWizard.Show
    MsgBox tf("enrollment.ribbon.package_next", "Следующий военнослужащий пакета подготовлен. OrderDraftId: {draftId}", "{draftId}", orderDraftId), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.package_next", "Ошибка подготовки следующей карточки пакета: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentExportSelectedClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim rowNum As Long
    Dim orderDraftId As String
    Dim exportScope As String
    Dim outputPath As String
    Dim exportErrorText As String

    rowNum = mdlEnrollmentWorkflow.ResolveActiveEnrollmentRow()
    orderDraftId = mdlEnrollmentOrderExport.GetOrderDraftIdForRow(rowNum)
    exportScope = mdlEnrollmentOrderExport.GetExportScopeText(orderDraftId, rowNum)
    exportErrorText = mdlEnrollmentOrderExport.GetEnrollmentExportBlockingIssues(orderDraftId, rowNum)
    If exportErrorText <> "" Then
        MsgBox tf("enrollment.form.message.export_blocked", "Word-приказ не сформирован.{nl}{error}", "{nl}", vbCrLf, "{error}", exportErrorText), vbExclamation, t("enrollment.caption.main", "Зачисление")
        Exit Sub
    End If
    outputPath = mdlEnrollmentOrderExport.ExportSelectedEnrollmentPackage()
    If mdlEnrollmentOrderExport.IsEnrollmentExportErrorResult(outputPath) Then
        MsgBox tf("enrollment.form.message.export_blocked", "Word-приказ не сформирован.{nl}{error}", "{nl}", vbCrLf, "{error}", mdlEnrollmentOrderExport.GetEnrollmentExportErrorText(outputPath)), vbExclamation, t("enrollment.caption.main", "Зачисление")
        Exit Sub
    End If
    MsgBox tf("enrollment.ribbon.exported", "Пакет приказа сформирован ({scope}).{nl}Файл: {path}", "{scope}", exportScope, "{nl}", vbCrLf, "{path}", outputPath), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.export_selected", "Ошибка экспорта пакета по выбранной строке зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentExportByDraftIdClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim orderDraftId As String
    Dim exportScope As String
    Dim outputPath As String
    Dim exportErrorText As String

    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    orderDraftId = Trim$(CStr(Application.InputBox( _
        Prompt:=t("enrollment.ribbon.prompt.order_draft_id", "Введите OrderDraftId:"), _
        Title:=t("enrollment.ribbon.caption.export_by_id", "Экспорт пакета зачисления"), _
        Type:=2)))

    If orderDraftId = "" Or orderDraftId = "False" Then
        MsgBox t("enrollment.ribbon.message.order_draft_id_empty", "OrderDraftId не указан. Экспорт отменён."), vbInformation, t("enrollment.caption.main", "Зачисление")
        Exit Sub
    End If

    exportScope = mdlEnrollmentOrderExport.GetExportScopeText(orderDraftId, 0)
    exportErrorText = mdlEnrollmentOrderExport.GetEnrollmentExportBlockingIssues(orderDraftId, 0)
    If exportErrorText <> "" Then
        MsgBox tf("enrollment.form.message.export_blocked", "Word-приказ не сформирован.{nl}{error}", "{nl}", vbCrLf, "{error}", exportErrorText), vbExclamation, t("enrollment.caption.main", "Зачисление")
        Exit Sub
    End If
    outputPath = mdlEnrollmentOrderExport.ExportEnrollmentOrderByDraftId(orderDraftId, 0)
    If mdlEnrollmentOrderExport.IsEnrollmentExportErrorResult(outputPath) Then
        MsgBox tf("enrollment.form.message.export_blocked", "Word-приказ не сформирован.{nl}{error}", "{nl}", vbCrLf, "{error}", mdlEnrollmentOrderExport.GetEnrollmentExportErrorText(outputPath)), vbExclamation, t("enrollment.caption.main", "Зачисление")
        Exit Sub
    End If
    MsgBox tf("enrollment.ribbon.exported", "Пакет приказа сформирован ({scope}).{nl}Файл: {path}", "{scope}", exportScope, "{nl}", vbCrLf, "{path}", outputPath), vbInformation, t("enrollment.caption.main", "Зачисление")
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.export_by_id", "Ошибка экспорта пакета по OrderDraftId: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentRefreshClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT).Activate
    mdlEnrollmentWorkflow.RefreshEnrollmentSuggestions
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.refresh", "Ошибка обновления предложений по зачислению: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentValidateClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT).Activate
    mdlEnrollmentWorkflow.ValidateEnrollmentSheet
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.validate", "Ошибка проверки журнала зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub OnEnrollmentTransferClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT).Activate
    mdlEnrollmentWorkflow.GeneratePaymentsFromEnrollmentSheet
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.transfer", "Ошибка переноса выплат из зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

' Handler for settings (Обновлено под новую систему лицензирования)
Sub ShowSettings(control As IRibbonControl)
    MsgBox BuildSettingsDiagnosticText(), vbInformation, t("ribbon.settings.title", "Настройки и проверка")
End Sub

Public Function BuildSettingsDiagnosticText() As String
    Dim settingsText As String

    settingsText = t("ribbon.settings.header", "=== НАСТРОЙКИ МАКРОСОВ ===") & vbCrLf & vbCrLf
    settingsText = settingsText & tf("ribbon.settings.current_folder", "[ПАПКА] Текущая папка: {path}", "{path}", ThisWorkbook.Path) & vbCrLf & vbCrLf
    settingsText = settingsText & t("ribbon.settings.templates_check", "[ПРОВЕРКА] Проверка шаблонов:") & vbCrLf
    settingsText = settingsText & BuildTemplateStatusLine("Шаблон_Справка.docx")
    settingsText = settingsText & BuildTemplateStatusLine("Шаблон_Рапорт.docx")

    settingsText = settingsText & vbCrLf & t("ribbon.settings.activation_prefix", "[СТАТУС АКТИВАЦИИ]: ")
    Select Case modActivation.GetLicenseStatus()
        Case 0
            settingsText = settingsText & tf("ribbon.settings.license.personal", "ПЕРСОНАЛЬНАЯ ЛИЦЕНЗИЯ (до {date})", "{date}", modActivation.GetLicenseExpiryDateStr()) & vbCrLf
        Case 3
            settingsText = settingsText & tf("ribbon.settings.license.corporate", "КОРПОРАТИВНАЯ ЛИЦЕНЗИЯ (до {date})", "{date}", modActivation.GetLicenseExpiryDateStr()) & vbCrLf
        Case 4
            settingsText = settingsText & tf("ribbon.settings.license.trial", "ОЗНАКОМИТЕЛЬНЫЙ ПЕРИОД (до {date})", "{date}", modActivation.GetLicenseExpiryDateStr()) & vbCrLf
        Case 1
            settingsText = settingsText & t("ribbon.settings.license.limited", "ОГРАНИЧЕННАЯ ВЕРСИЯ (срок истек)") & vbCrLf
        Case 2
            settingsText = settingsText & t("ribbon.settings.license.blocked_time", "БЛОКИРОВКА (сбой системного времени)") & vbCrLf
    End Select

    settingsText = settingsText & vbCrLf & tf("ribbon.settings.version", "[ВЕРСИЯ] Версия макросов: {version}", "{version}", modActivation.PRODUCT_VERSION)
    BuildSettingsDiagnosticText = settingsText
End Function

Private Function BuildTemplateStatusLine(ByVal templateFileName As String) As String
    If dir(ThisWorkbook.Path & "\" & templateFileName) <> "" Then
        BuildTemplateStatusLine = tf("ribbon.settings.template_found", "[+] {template} - найден", "{template}", templateFileName) & vbCrLf
    Else
        BuildTemplateStatusLine = tf("ribbon.settings.template_missing", "[-] {template} - НЕ НАЙДЕН", "{template}", templateFileName) & vbCrLf
    End If
End Function

' Handler for "Remove Duplicate Modules" button
Public Sub OnRemoveDuplicateModulesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call MdlBackup.RemoveDuplicateModules
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.remove_duplicates", "Ошибка при удалении дубликатов: {error}", "{error}", Err.description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub RunValidateZP12(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    mdlZP12Validation.ValidateZP12Template
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox tf("ribbon.error.zp12_validation", "Ошибка проверки Д89: {error}", "{error}", Err.description), vbCritical, t("ribbon.zp12.title", "Проверка Д89")
End Sub

