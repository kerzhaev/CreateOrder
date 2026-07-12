VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEnrollmentWizard
   Caption         =   "UserForm1"
   ClientHeight    =   8925.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   OleObjectBlob   =   "frmEnrollmentWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEnrollmentWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RESULT_COL_PERSONAL_NUMBER As Long = 0
Private Const RESULT_COL_FIO As Long = 1
Private Const RESULT_COL_RANK As Long = 2
Private Const RESULT_COL_POSITION As Long = 3
Private Const RESULT_COL_SECTION As Long = 4

Private mpWizard As Object
Private pgEmployee As Object
Private pgDocs As Object
Private pgMonthly As Object
Private pgOneTime As Object
Private pgAdvanced As Object
Private pgExtras As Object
Private pgPreview As Object

Private txtEmployeeFIO As Object
Private txtEmployeeNumber As Object
Private txtEmployeeTableNumber As Object
Private txtEmployeeRank As Object
Private txtEmployeeServiceCategory As Object
Private txtEmployeeContractKind As Object
Private txtEmployeeContractBasis As Object
Private txtEmployeeVus As Object
Private txtEmployeePosition As Object
Private txtEmployeeSection As Object
Private txtEmployeeMilitaryUnit As Object
Private txtEmployeeTariff As Object
Private txtEmployeePositionSalary As Object
Private txtEmployeeRankSalary As Object

Private txtOrderDate As Object
Private txtOrderDraftId As Object
Private txtOrderNumber As Object
Private txtOrderIssuer As Object
Private txtArrivalSource As Object
Private txtPrescriptionNumber As Object
Private txtPrescriptionDate As Object
Private txtReportNumber As Object
Private txtReportDate As Object
Private txtReportInfo As Object
Private txtAssignmentInfo As Object
Private txtAcceptDate As Object
Private txtEnrollDate As Object
Private txtDutyStartDate As Object
Private txtManualStart As Object
Private txtStandardStart As Object
Private txtPreferentialStart As Object
Private txtBasisSection1 As Object
Private txtBasisSection2 As Object

Private chkStdDuty As Object
Private txtStdDutyPercent As Object
Private chkStdSpecial As Object
Private txtStdSpecialPercent As Object
Private chkStdTariff As Object
Private txtStdTariffPercent As Object
Private chkStdContract430 As Object
Private txtStdContract430Percent As Object
Private chkPremium As Object
Private txtPremiumPercent As Object
Private txtPremiumStart As Object
Private txtPremiumEnd As Object
Private txtClassParam As Object
Private chkClass As Object
Private txtClassPercent As Object
Private txtFizoParam As Object
Private chkFizo As Object
Private txtFizoPercent As Object
Private txtSecrecyParam As Object
Private chkSecrecy As Object
Private txtSecrecyPercent As Object
Private txtAchievementParam As Object
Private chkAchievement As Object
Private txtAchievementAmount As Object
Private chkPreferential As Object
Private txtPreferentialCoeff As Object

Private chkLift As Object
Private txtLiftAmount As Object
Private txtLiftDate As Object
Private chkPerDiem As Object
Private txtPerDiemDays As Object
Private txtPerDiemAmount As Object
Private txtPerDiemDate As Object
Private chkEdv As Object
Private txtEdvAmount As Object
Private txtEdvDate As Object
Private txtBirthDate As Object
Private txtBirthPlace As Object
Private txtCitizenship As Object
Private txtInn As Object
Private txtSnils As Object
Private txtPassportSeries As Object
Private txtPassportNumber As Object
Private txtPassportIssuer As Object
Private txtPassportIssueDate As Object
Private txtPassportCode As Object
Private txtBankAccount As Object
Private txtBankName As Object
Private txtRequisitesNote As Object

Private txtPreferentialBasis As Object
Private txtPremiumBasis As Object
Private txtLiftBasis As Object
Private txtPerDiemBasis As Object
Private txtEdvBasis As Object
Private txtClassBasis As Object
Private txtFizoBasis As Object
Private txtSecrecyBasis As Object
Private txtAchievementBasis As Object
Private txtStdDutyBasis As Object
Private txtStdSpecialBasis As Object
Private txtStdTariffBasis As Object
Private txtStdContract430Basis As Object

Private txtExtraMonthlyName(1 To 4) As Object
Private txtExtraMonthlyParam(1 To 4) As Object
Private txtExtraMonthlyAmount(1 To 4) As Object
Private txtExtraMonthlyStart(1 To 4) As Object
Private txtExtraMonthlyBasis(1 To 4) As Object
Private chkExtraMonthly(1 To 4) As Object

Private txtExtraOneTimeName(1 To 3) As Object
Private txtExtraOneTimeAmount(1 To 3) As Object
Private txtExtraOneTimeDate(1 To 3) As Object
Private txtExtraOneTimeBasis(1 To 3) As Object
Private chkExtraOneTime(1 To 3) As Object

Private txtPreviewStatus As Object
Private txtPreviewReady As Object
Private txtPreviewIssues As Object
Private txtPreviewStandard As Object
Private txtPreviewPersonal As Object
Private txtPreviewSection1 As Object
Private txtPreviewSection2 As Object

Private WithEvents btnSaveGenerateDynamic As MSForms.CommandButton
Private WithEvents btnSaveContinueDynamic As MSForms.CommandButton
Private WithEvents btnLoadFromInlineSearchDynamic As MSForms.CommandButton
Private WithEvents btnCheckDynamic As MSForms.CommandButton
Private WithEvents btnSaveCardDynamic As MSForms.CommandButton
Private WithEvents btnExportPackageDynamic As MSForms.CommandButton

Private currentSourceMode As String
Private Const PREVIEW_PAGE_INDEX As Long = 6

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    mdlHelper.EnsureStaffColumnsInitialized

    HideLegacyControls
    ConfigureSearchArea
    EnsureDynamicActionButtons
    ConfigureWindow
    ConfigureButtons
    CreateWizardUi
    currentSourceMode = "manual"
    ReloadFromBackend
    lblStatus.Caption = t("enrollment.form.status.ready_to_pick", "Выберите сотрудника из листа 'Штат' или заполните карточку вручную. После выбора проверьте страницы мастера.")
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "frmEnrollmentWizard.UserForm_Initialize", Err.Description
End Sub

Private Sub txtSearch_Change()
    RefreshInlineSearchResults
End Sub

Private Sub RefreshInlineSearchResults()
    On Error GoTo ErrorHandler

    Dim wsStaff As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim foundCount As Long
    Dim query As String
    Dim colTableNumber As Long
    Dim fioValue As String
    Dim lnValue As String
    Dim tableValue As String

    Set wsStaff = mdlHelper.GetStaffWorksheet()
    If wsStaff Is Nothing Then Exit Sub

    query = LCase$(Trim$(txtSearch.Text))
    lstResults.Clear

    If Len(query) < 2 Then
        lblStatus.Caption = t("common.status_enter_min_chars", "Введите не менее 2 символов.")
        If Not btnLoadFromInlineSearchDynamic Is Nothing Then btnLoadFromInlineSearchDynamic.Enabled = False
        Exit Sub
    End If

    lastRow = wsStaff.Cells(wsStaff.Rows.Count, mdlHelper.colFIO_Global).End(xlUp).Row
    colTableNumber = mdlHelper.FindTableNumberColumn(wsStaff)

    For rowNum = 2 To lastRow
        fioValue = LCase$(Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colFIO_Global).Value)))
        lnValue = LCase$(Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colLichniyNomer_Global).Value)))
        tableValue = ""
        If colTableNumber > 0 Then tableValue = LCase$(Trim$(CStr(wsStaff.Cells(rowNum, colTableNumber).Value)))

        If InStr(1, fioValue, query, vbTextCompare) > 0 _
            Or InStr(1, lnValue, query, vbTextCompare) > 0 _
            Or (tableValue <> "" And InStr(1, tableValue, query, vbTextCompare) > 0) Then
            AddSearchResult wsStaff, rowNum, foundCount
            foundCount = foundCount + 1
        End If
    Next rowNum

    If foundCount = 0 Then
        lblStatus.Caption = t("common.status_none", "Совпадения не найдены.")
        If Not btnLoadFromInlineSearchDynamic Is Nothing Then btnLoadFromInlineSearchDynamic.Enabled = False
    Else
        lblStatus.Caption = tf("common.status_found", "Найдено: {count}", "{count}", foundCount)
        If foundCount = 1 Then lstResults.ListIndex = 0
        If Not btnLoadFromInlineSearchDynamic Is Nothing Then btnLoadFromInlineSearchDynamic.Enabled = True
    End If
    Exit Sub

ErrorHandler:
    lblStatus.Caption = tf("enrollment.form.error.search", "Ошибка поиска сотрудника: {error}", "{error}", Err.Description)
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstResults.ListCount > 0 Then
            lstResults.SetFocus
            If lstResults.ListIndex < 0 Then lstResults.ListIndex = 0
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        If lstResults.ListCount = 1 Then
            lstResults.ListIndex = 0
            btnSelect_Click
        End If
        KeyCode = 0
    End If
End Sub

Private Sub lstResults_Click()
    If lstResults.ListCount = 0 Or lstResults.ListIndex < 0 Then Exit Sub
    lblStatus.Caption = tf("enrollment.form.status.preview_found", "Найден сотрудник: {fio}", "{fio}", CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_FIO)))
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnSelect_Click
End Sub

Private Sub lstResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnSelect_Click
        KeyCode = 0
    End If
End Sub

Private Sub btnSelect_Click()
    On Error GoTo ErrorHandler

    Dim selectedNumber As String

    selectedNumber = PickEmployeeFromStaff()
    If selectedNumber = "" Then Exit Sub

    LoadEmployeeFromStaffNumber selectedNumber
    mpWizard.Value = 0
    Exit Sub

ErrorHandler:
    MsgBox tf("enrollment.form.error.load_employee", "Ошибка загрузки сотрудника: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Sub LoadEmployeeFromStaffNumber(ByVal employeeNumber As String)
    Dim wsStaff As Worksheet
    Dim staffRow As Long
    Dim tableColumn As Long
    Dim staffData As Object

    mdlHelper.EnsureStaffColumnsInitialized
    Set wsStaff = mdlHelper.GetStaffWorksheet()
    If wsStaff Is Nothing Then
        Err.Raise vbObjectError + 1810, "frmEnrollmentWizard.LoadEmployeeFromStaffNumber", t("form.select_employee.message.staff_columns_error", "Не удалось определить обязательные столбцы листа 'Штат'.")
    End If

    staffRow = ResolveStaffRowByAnyNumber(wsStaff, employeeNumber)
    If staffRow < 2 Then
        Err.Raise vbObjectError + 1811, "frmEnrollmentWizard.LoadEmployeeFromStaffNumber", tf("enrollment.form.error.employee_not_found", "Сотрудник с номером {number} не найден на листе 'Штат'.", "{number}", employeeNumber)
    End If

    tableColumn = mdlHelper.FindTableNumberColumn(wsStaff)
    Set staffData = mdlHelper.GetStaffData(employeeNumber, True)
    currentSourceMode = "staff"
    txtEmployeeNumber.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colLichniyNomer_Global).Value))
    If tableColumn > 0 Then
        txtEmployeeTableNumber.Value = Trim$(CStr(wsStaff.Cells(staffRow, tableColumn).Value))
    Else
        txtEmployeeTableNumber.Value = ""
    End If
    txtEmployeeFIO.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colFIO_Global).Value))
    txtEmployeeRank.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colZvanie_Global).Value))
    txtEmployeePosition.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colDolzhnost_Global).Value))
    txtEmployeeSection.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colVoinskayaChast_Global).Value))
    txtEmployeeMilitaryUnit.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colVoinskayaChast_Global).Value))
    txtEmployeeServiceCategory.Value = StaffDictValue(staffData, mdlHelper.Ru(1043, 1088, 1091, 1087, 1087, 1072, 32, 1089, 1086, 1090, 1088, 1091, 1076, 1085, 1080, 1082, 1086, 1074))
    txtEmployeeContractKind.Value = StaffDictValue(staffData, mdlHelper.Ru(1042, 1080, 1076, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072))
    txtEmployeeVus.Value = StaffDictValue(staffData, mdlHelper.Ru(1042, 1059, 1057))
    txtEmployeeTariff.Value = StaffDictValue(staffData, mdlHelper.Ru(1058, 1072, 1088, 1080, 1092, 1085, 1099, 1081, 32, 1088, 1072, 1079, 1088, 1103, 1076))
    txtBirthDate.Value = StaffDictDateValue(staffData, mdlHelper.Ru(1044, 1072, 1090, 1072, 32, 1088, 1086, 1078, 1076, 1077, 1085, 1080, 1103))
    txtCitizenship.Value = StaffDictValue(staffData, mdlHelper.Ru(1043, 1088, 1072, 1078, 1076, 1072, 1085, 1089, 1090, 1074, 1086))
    txtBankAccount.Value = StaffDictValue(staffData, mdlHelper.Ru(1053, 1086, 1084, 1077, 1088, 32, 1089, 1095, 1077, 1090, 1072, 32, 1074, 32, 1073, 1072, 1085, 1082, 1077))

    PushFormToBackend
    mdlEnrollmentWorkflow.RefreshEnrollmentForm
    ReloadFromBackend

    lblStatus.Caption = tf("enrollment.form.status.employee_loaded", "Данные из листа 'Штат' загружены. Сотрудник: {fio}", "{fio}", txtEmployeeFIO.Value)
End Sub

Private Function StaffDictValue(ByVal staffData As Object, ByVal key As String) As String
    If staffData Is Nothing Then Exit Function
    If Not staffData.Exists(key) Then Exit Function

    StaffDictValue = Trim$(CStr(staffData(key)))
End Function

Private Function StaffDictDateValue(ByVal staffData As Object, ByVal key As String) As String
    Dim rawValue As Variant

    If staffData Is Nothing Then Exit Function
    If Not staffData.Exists(key) Then Exit Function

    rawValue = staffData(key)
    If IsDate(rawValue) Then
        StaffDictDateValue = Format$(CDate(rawValue), "dd.mm.yyyy")
    Else
        StaffDictDateValue = Trim$(CStr(rawValue))
    End If
End Function

Public Function GetEmployeeSnapshot() As String
    GetEmployeeSnapshot = currentSourceMode & "|" & _
        Trim$(CStr(txtEmployeeFIO.Value)) & "|" & _
        Trim$(CStr(txtEmployeeNumber.Value)) & "|" & _
        Trim$(CStr(txtEmployeePosition.Value)) & "|" & _
        Trim$(CStr(txtEmployeeMilitaryUnit.Value)) & "|" & _
        Trim$(CStr(txtEmployeeTableNumber.Value)) & "|" & _
        Trim$(CStr(txtEmployeeServiceCategory.Value)) & "|" & _
        Trim$(CStr(txtEmployeeContractKind.Value)) & "|" & _
        Trim$(CStr(txtEmployeeVus.Value)) & "|" & _
        Trim$(CStr(txtEmployeeTariff.Value)) & "|" & _
        Trim$(CStr(txtBirthDate.Value)) & "|" & _
        Trim$(CStr(txtCitizenship.Value)) & "|" & _
        Trim$(CStr(txtBankAccount.Value))
End Function

Public Function ProbeInlineSearch(ByVal queryText As String) As String
    txtSearch.Text = queryText
    RefreshInlineSearchResults

    ProbeInlineSearch = CStr(lstResults.ListCount)
    If lstResults.ListCount > 0 Then
        ProbeInlineSearch = ProbeInlineSearch & "|" & CStr(lstResults.List(0, RESULT_COL_FIO)) & "|" & CStr(lstResults.List(0, RESULT_COL_PERSONAL_NUMBER))
    End If
End Function

Private Function PickEmployeeFromStaff() As String
    frmSelectEmployee.selectedLichniyNomer = ""
    frmSelectEmployee.selectedFIO = ""
    frmSelectEmployee.isCancelled = True
    frmSelectEmployee.Show

    If frmSelectEmployee.isCancelled Then Exit Function
    PickEmployeeFromStaff = Trim$(frmSelectEmployee.selectedLichniyNomer)
End Function

Private Function ResolveStaffRowByAnyNumber(ByVal wsStaff As Worksheet, ByVal employeeNumber As String) As Long
    Dim tableColumn As Long

    ResolveStaffRowByAnyNumber = mdlHelper.FindStaffRow(wsStaff, Trim$(employeeNumber), mdlHelper.colLichniyNomer_Global)
    If ResolveStaffRowByAnyNumber >= 2 Then Exit Function

    tableColumn = mdlHelper.FindTableNumberColumn(wsStaff)
    If tableColumn > 0 Then
        ResolveStaffRowByAnyNumber = mdlHelper.FindStaffRow(wsStaff, Trim$(employeeNumber), tableColumn)
    End If
End Function

Private Sub btnLoadFromInlineSearchDynamic_Click()
    If lstResults.ListCount = 0 Or lstResults.ListIndex < 0 Then
        MsgBox t("form.select_employee.message.choose_from_list", "Выберите сотрудника из списка."), vbExclamation, t("common.attention", "Внимание")
        Exit Sub
    End If

    LoadEmployeeFromStaffNumber CStr(lstResults.List(lstResults.ListIndex, RESULT_COL_PERSONAL_NUMBER))
    mpWizard.Value = 0
End Sub

Private Sub btnAddPeriod_Click()
    On Error GoTo ErrorHandler
    PerformCheckPreview True
    Exit Sub
ErrorHandler:
    MsgBox tf("enrollment.form.error.refresh", "Ошибка обновления карточки зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub
Private Sub btnEditPeriod_Click()
    On Error GoTo ErrorHandler

    PerformSaveCard True
    Exit Sub

ErrorHandler:
    MsgBox tf("enrollment.form.error.save", "Ошибка сохранения карточки зачисления: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Private Sub btnDeletePeriod_Click()
    On Error GoTo ErrorHandler

    PerformExportPackage True
    Exit Sub

ErrorHandler:
    MsgBox tf("enrollment.form.error.export", "Ошибка экспорта приказа о зачислении: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Private Sub btnCheckDynamic_Click()
    btnAddPeriod_Click
End Sub

Private Sub btnSaveCardDynamic_Click()
    btnEditPeriod_Click
End Sub

Private Sub btnExportPackageDynamic_Click()
    btnDeletePeriod_Click
End Sub

Private Sub btnSaveGenerateDynamic_Click()
    On Error GoTo ErrorHandler

    PerformSaveGenerate False, True
    Exit Sub

ErrorHandler:
    MsgBox tf("enrollment.form.error.save_generate", "Ошибка сохранения карточки зачисления с подготовкой выплат: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Private Sub btnSaveContinueDynamic_Click()
    On Error GoTo ErrorHandler

    PerformSaveContinuePackage True
    Exit Sub

ErrorHandler:
    MsgBox tf("enrollment.ribbon.error.package_next", "Ошибка подготовки следующей карточки пакета: {error}", "{error}", Err.Description), vbCritical, t("common.error", "Ошибка")
End Sub

Public Function RunSaveGenerateAction() As String
    RunSaveGenerateAction = PerformSaveGenerate(False, False)
End Function

Public Function RunSaveContinuePackageAction() As String
    RunSaveContinuePackageAction = PerformSaveContinuePackage(False)
End Function

Public Function RunSaveCardAction() As String
    RunSaveCardAction = PerformSaveCard(False)
End Function

Public Function RunCheckAction() As String
    RunCheckAction = PerformCheckPreview(True)
End Function
Public Function RunExportAction() As String
    RunExportAction = PerformExportPackage(False)
End Function

Public Function ProbeLayoutSnapshot() As String
    ProbeLayoutSnapshot = CStr(CLng(Me.Height)) & "|" & _
        CStr(CLng(Me.Width)) & "|" & _
        CStr(CLng(mpWizard.Height)) & "|" & _
        CStr(CLng(mpWizard.Width)) & "|" & _
        CStr(CLng(btnCheckDynamic.Top)) & "|" & _
        CStr(CLng(btnExportPackageDynamic.Top)) & "|" & _
        CStr(CLng(btnClose.Top + btnClose.Height)) & "|" & _
        CStr(CLng(btnSelect.Left + btnSelect.Width)) & "|" & _
        CStr(CLng(chkPreferential.Top)) & "|" & _
        CStr(CLng(chkStdDuty.Top)) & "|" & _
        CStr(CLng(txtRequisitesNote.Left + txtRequisitesNote.Width)) & "|" & _
        CStr(CLng(txtRequisitesNote.Top + txtRequisitesNote.Height)) & "|" & _
        CStr(btnLoadFromInlineSearchDynamic.Caption) & "|" & _
        CStr(CLng(txtExtraMonthlyBasis(1).Top + txtExtraMonthlyBasis(1).Height)) & "|" & _
        CStr(CLng(txtExtraMonthlyName(2).Top)) & "|" & _
        CStr(CLng(txtExtraMonthlyBasis(4).Top + txtExtraMonthlyBasis(4).Height)) & "|" & _
        CStr(CLng(txtExtraOneTimeName(1).Top)) & "|" & _
        CStr(CLng(txtExtraOneTimeBasis(3).Top + txtExtraOneTimeBasis(3).Height)) & "|" & _
        CStr(CLng(pgExtras.ScrollHeight)) & "|" & _
        CStr(CLng(txtExtraMonthlyName(1).Top + txtExtraMonthlyName(1).Height)) & "|" & _
        CStr(CLng(txtExtraMonthlyBasis(1).Top)) & "|" & _
        CStr(CLng(txtExtraOneTimeName(1).Top + txtExtraOneTimeName(1).Height)) & "|" & _
        CStr(CLng(txtExtraOneTimeBasis(1).Top)) & "|" & _
        tf("enrollment.field.extra_monthly_name_short", "Ежемес. #{index}: вид", "{index}", 1) & "|" & _
        tf("enrollment.field.extra_onetime_name_short", "Разовая #{index}: вид", "{index}", 1)
End Function

Public Function ProbeFullCardSnapshot() As String
    ProbeFullCardSnapshot = SafeText(txtEmployeeFIO.Value) & "|" & _
        SafeText(txtEmployeeNumber.Value) & "|" & _
        SafeText(txtOrderDraftId.Value) & "|" & _
        SafeText(txtPremiumEnd.Value) & "|" & _
        SafeText(txtPassportSeries.Value) & "|" & _
        SafeText(txtPassportNumber.Value) & "|" & _
        SafeText(txtBankAccount.Value) & "|" & _
        SafeText(txtBankName.Value) & "|" & _
        SafeText(txtBasisSection2.Value) & "|" & _
        SafeText(txtEdvBasis.Value) & "|" & _
        SafeText(txtExtraMonthlyName(1).Value) & "|" & _
        SafeText(txtExtraMonthlyAmount(1).Value) & "|" & _
        SafeText(txtExtraMonthlyBasis(1).Value) & "|" & _
        SafeText(txtExtraOneTimeName(1).Value) & "|" & _
        SafeText(txtExtraOneTimeAmount(1).Value) & "|" & _
        SafeText(txtExtraOneTimeBasis(1).Value)
End Function

Private Function PerformCheckPreview(Optional ByVal keepPreviewPage As Boolean = True) As String
    PushFormToBackend
    mdlEnrollmentWorkflow.RefreshEnrollmentForm
    ReloadFromBackend
    If keepPreviewPage Then mpWizard.Value = PREVIEW_PAGE_INDEX
    lblStatus.Caption = t("enrollment.form.status.checked", "Карточка зачисления проверена. Открыта вкладка предпросмотра.")
    PerformCheckPreview = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_word_ready")) & "|" & _
        CStr(Len(SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_section1")))) & "|" & _
        CStr(Len(SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_section2")))) & "|" & _
        CStr(Len(SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_issues"))))
End Function

Private Function PerformSaveCard(Optional ByVal showMessage As Boolean = True) As String
    Dim targetRow As Long
    Dim orderDraftId As String

    PushFormToBackend
    targetRow = mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet(False)
    orderDraftId = mdlEnrollmentOrderExport.GetOrderDraftIdForRow(targetRow)
    ReloadFromBackend
    lblStatus.Caption = tf("enrollment.form.status.saved", "Карточка сохранена. OrderDraftId: {draftId}.", "{draftId}", orderDraftId)

    If showMessage Then
        MsgBox tf("enrollment.form.message.saved", "Карточка зачисления сохранена в строку {row}. OrderDraftId: {draftId}.", "{row}", targetRow, "{draftId}", orderDraftId), vbInformation, t("enrollment.caption.main", "Зачисление")
    End If

    PerformSaveCard = CStr(targetRow) & "|" & orderDraftId
End Function

Private Function PerformSaveGenerate(Optional ByVal keepPreviewPage As Boolean = False, Optional ByVal showMessage As Boolean = True) As String
    Dim targetRow As Long
    Dim createdCount As Long
    Dim orderDraftId As String

    PushFormToBackend
    targetRow = mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet(False)
    createdCount = mdlEnrollmentWorkflow.GeneratePaymentsFromEnrollmentRowDirect(targetRow)
    orderDraftId = mdlEnrollmentOrderExport.GetOrderDraftIdForRow(targetRow)
    ReloadFromBackend
    If keepPreviewPage Then mpWizard.Value = PREVIEW_PAGE_INDEX
    lblStatus.Caption = tf("enrollment.form.status.saved_generated", "Карточка сохранена. Подготовлено выплат: {count}.", "{count}", createdCount)

    If showMessage Then
        MsgBox tf("enrollment.form.message.saved_generated", "Карточка сохранена в строку {row}. OrderDraftId: {draftId}.{nl}Подготовлено выплат: {count}.", "{row}", targetRow, "{draftId}", orderDraftId, "{nl}", vbCrLf, "{count}", createdCount), vbInformation, t("enrollment.caption.main", "Зачисление")
    End If

    PerformSaveGenerate = CStr(targetRow) & "|" & orderDraftId & "|" & CStr(createdCount)
End Function

Private Function PerformExportPackage(Optional ByVal showMessage As Boolean = True) As String
    Dim targetRow As Long
    Dim orderDraftId As String
    Dim exportScope As String
    Dim outputPath As String
    Dim exportErrorText As String

    On Error GoTo ExportBlocked

    PushFormToBackend
    targetRow = mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet(False)
    orderDraftId = mdlEnrollmentOrderExport.GetOrderDraftIdForRow(targetRow)
    exportScope = mdlEnrollmentOrderExport.GetExportScopeText(orderDraftId, targetRow)
    exportErrorText = mdlEnrollmentOrderExport.GetEnrollmentExportBlockingIssues(orderDraftId, targetRow)
    If exportErrorText <> "" Then GoTo ExportBlocked

    outputPath = mdlEnrollmentOrderExport.ExportEnrollmentOrderByDraftId(orderDraftId, targetRow)
    If mdlEnrollmentOrderExport.IsEnrollmentExportErrorResult(outputPath) Then
        exportErrorText = mdlEnrollmentOrderExport.GetEnrollmentExportErrorText(outputPath)
        GoTo ExportBlocked
    End If

    ReloadFromBackend
    mpWizard.Value = PREVIEW_PAGE_INDEX

    lblStatus.Caption = tf("enrollment.form.status.exported", "Сформирован пакет приказа: {scope}.", "{scope}", exportScope)
    If showMessage Then
        MsgBox tf("enrollment.form.message.exported", "Пакет приказа сформирован ({scope}).{nl}Файл: {path}", "{scope}", exportScope, "{nl}", vbCrLf, "{path}", outputPath), vbInformation, t("enrollment.caption.main", "Зачисление")
    End If

    PerformExportPackage = outputPath
    Exit Function

ExportBlocked:
    If exportErrorText = "" Then exportErrorText = Err.Description
    On Error Resume Next
    ReloadFromBackend
    mpWizard.Value = PREVIEW_PAGE_INDEX
    lblStatus.Caption = t("enrollment.form.status.export_blocked", "Экспорт заблокирован. Проверьте замечания на вкладке предпросмотра.")
    On Error GoTo 0

    If showMessage Then
        MsgBox tf("enrollment.form.message.export_blocked", "Word-приказ не сформирован.{nl}{error}", "{nl}", vbCrLf, "{error}", exportErrorText), vbExclamation, t("enrollment.caption.main", "Зачисление")
    End If

    PerformExportPackage = "ERROR: " & exportErrorText
End Function

Private Function PerformSaveContinuePackage(Optional ByVal showMessage As Boolean = True) As String
    Dim orderDraftId As String

    PushFormToBackend
    orderDraftId = mdlEnrollmentWorkflow.SaveEnrollmentFormAndContinuePackage()
    ReloadFromBackend
    mpWizard.Value = 0
    lblStatus.Caption = tf("enrollment.form.status.package_next", "Подготовлена новая карточка в пакете {draftId}. Заполните сведения о следующем военнослужащем.", "{draftId}", orderDraftId)

    If showMessage Then
        MsgBox tf("enrollment.form.message.package_next", "Следующий военнослужащий пакета подготовлен. OrderDraftId: {draftId}", "{draftId}", orderDraftId), vbInformation, t("enrollment.caption.main", "Зачисление")
    End If

    PerformSaveContinuePackage = orderDraftId & "|" & SafeText(mdlEnrollmentWorkflow.GetBackendValue("fio")) & "|" & SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_number")) & "|" & SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_duty_enabled"))
End Function

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub ConfigureWindow()
    Me.Caption = t("enrollment.form.title", "Мастер зачисления")
    Me.Width = 880
    Me.Height = 720

    lblStatus.Left = 18
    lblStatus.Top = 22
    lblStatus.Width = 680
    lblStatus.Height = 36

    ConfigureInlineSearchUi
End Sub

Private Sub HideLegacyControls()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    lstPeriods.Visible = False
    txtPeriodStart.Visible = False
    txtPeriodEnd.Visible = False
    cmbReason.Visible = False
    Label_PeriodStart.Visible = False
    Label_PeriodEnd.Visible = False
    Label_Reason.Visible = False
    lblFIO.Visible = False
    lblZvanie.Visible = False
    lblDolzhnost.Visible = False
    lblChast.Visible = False
End Sub

Private Sub ConfigureSearchArea()
    With lstResults
        .ColumnCount = 5
        .ColumnHeads = False
        .BoundColumn = 1
        .ColumnWidths = "70 pt;140 pt;85 pt;160 pt;140 pt"
        .IntegralHeight = False
        .ListStyle = fmListStylePlain
        .MultiSelect = fmMultiSelectSingle
        .Clear
    End With
End Sub

Private Sub ConfigureInlineSearchUi()
    txtSearch.Left = 18
    txtSearch.Top = 66
    txtSearch.Width = 280
    txtSearch.Height = 22
    txtSearch.Visible = True
    txtSearch.ControlTipText = t("enrollment.form.search.tip", "Введите ФИО, личный или табельный номер.")

    btnLoadFromInlineSearchDynamic.Caption = t("enrollment.form.button.load_from_search", "Загрузить из поиска")
    btnLoadFromInlineSearchDynamic.Left = 310
    btnLoadFromInlineSearchDynamic.Top = 64
    btnLoadFromInlineSearchDynamic.Width = 160
    btnLoadFromInlineSearchDynamic.Height = 26
    btnLoadFromInlineSearchDynamic.Visible = True
    btnLoadFromInlineSearchDynamic.Enabled = False

    lstResults.Left = 18
    lstResults.Top = 98
    lstResults.Width = 610
    lstResults.Height = 88
    lstResults.Visible = True
End Sub

Private Sub ConfigureButtons()
    EnsureDynamicActionButtons

    btnSelect.Caption = t("enrollment.form.button.pick_from_staff", "Выбрать сотрудника из штата")
    btnAddPeriod.Caption = t("enrollment.form.button.check", "Проверить и показать")
    btnEditPeriod.Caption = t("enrollment.form.button.save", "Сохранить карточку")
    btnDeletePeriod.Caption = t("enrollment.form.button.export", "Сохранить и экспортировать пакет")
    btnCheckDynamic.Caption = btnAddPeriod.Caption
    btnSaveCardDynamic.Caption = btnEditPeriod.Caption
    btnExportPackageDynamic.Caption = btnDeletePeriod.Caption
    btnSaveGenerateDynamic.Caption = t("enrollment.form.button.save_generate", "Сохранить и подготовить выплаты")
    btnSaveContinueDynamic.Caption = t("enrollment.form.button.save_continue_package", "Сохранить и следующий в пакете")
    btnClose.Caption = t("common.close", "Закрыть")

    On Error Resume Next
    btnAddPeriod.Visible = False
    btnEditPeriod.Visible = False
    btnDeletePeriod.Visible = False
    On Error GoTo 0

    btnSelect.Left = 650
    btnSelect.Top = 64
    btnSelect.Width = 195

    btnCheckDynamic.Left = 18
    btnCheckDynamic.Top = 620
    btnCheckDynamic.Width = 130
    btnCheckDynamic.Height = 28

    btnSaveCardDynamic.Left = 158
    btnSaveCardDynamic.Top = 620
    btnSaveCardDynamic.Width = 135
    btnSaveCardDynamic.Height = 28

    btnSaveGenerateDynamic.Left = 303
    btnSaveGenerateDynamic.Top = 620
    btnSaveGenerateDynamic.Width = 185
    btnSaveGenerateDynamic.Height = 28

    btnSaveContinueDynamic.Left = 498
    btnSaveContinueDynamic.Top = 620
    btnSaveContinueDynamic.Width = 200
    btnSaveContinueDynamic.Height = 28

    btnExportPackageDynamic.Left = 18
    btnExportPackageDynamic.Top = 654
    btnExportPackageDynamic.Width = 220
    btnExportPackageDynamic.Height = 28

    btnClose.Left = 254
    btnClose.Top = 654
    btnClose.Width = 64
End Sub

Private Sub EnsureDynamicActionButtons()
    If btnSaveGenerateDynamic Is Nothing Then
        Set btnSaveGenerateDynamic = Me.Controls.Add("Forms.CommandButton.1", "btnSaveGenerateDynamic", True)
        btnSaveGenerateDynamic.TakeFocusOnClick = False
    End If

    If btnSaveContinueDynamic Is Nothing Then
        Set btnSaveContinueDynamic = Me.Controls.Add("Forms.CommandButton.1", "btnSaveContinueDynamic", True)
        btnSaveContinueDynamic.TakeFocusOnClick = False
    End If

    If btnLoadFromInlineSearchDynamic Is Nothing Then
        Set btnLoadFromInlineSearchDynamic = Me.Controls.Add("Forms.CommandButton.1", "btnLoadFromInlineSearchDynamic", True)
        btnLoadFromInlineSearchDynamic.TakeFocusOnClick = False
    End If

    If btnCheckDynamic Is Nothing Then
        Set btnCheckDynamic = Me.Controls.Add("Forms.CommandButton.1", "btnCheckDynamic", True)
        btnCheckDynamic.TakeFocusOnClick = False
    End If

    If btnSaveCardDynamic Is Nothing Then
        Set btnSaveCardDynamic = Me.Controls.Add("Forms.CommandButton.1", "btnSaveCardDynamic", True)
        btnSaveCardDynamic.TakeFocusOnClick = False
    End If

    If btnExportPackageDynamic Is Nothing Then
        Set btnExportPackageDynamic = Me.Controls.Add("Forms.CommandButton.1", "btnExportPackageDynamic", True)
        btnExportPackageDynamic.TakeFocusOnClick = False
    End If
End Sub

Private Sub CreateWizardUi()
    Set mpWizard = Me.Controls.Add("Forms.MultiPage.1", "mpWizard", True)
    With mpWizard
        .Left = 18
        .Top = 196
        .Width = 830
        .Height = 410
    End With

    RemoveDefaultWizardPages

    Set pgEmployee = mpWizard.Pages(0)
    pgEmployee.Caption = t("enrollment.page.employee", "Военнослужащий")
    Set pgDocs = mpWizard.Pages.Add
    pgDocs.Caption = t("enrollment.page.docs", "Документы и даты")
    ConfigureScrollablePage pgDocs, 520
    Set pgMonthly = mpWizard.Pages.Add
    pgMonthly.Caption = t("enrollment.page.monthly", "Ежемесячные выплаты")
    Set pgOneTime = mpWizard.Pages.Add
    pgOneTime.Caption = t("enrollment.page.onetime", "Разовые выплаты и реквизиты")
    Set pgAdvanced = mpWizard.Pages.Add
    pgAdvanced.Caption = t("enrollment.page.advanced", "Основания и пакет")
    Set pgExtras = mpWizard.Pages.Add
    pgExtras.Caption = t("enrollment.page.extras", "Дополнительные выплаты")
    ConfigureScrollablePage pgExtras, 720
    Set pgPreview = mpWizard.Pages.Add
    pgPreview.Caption = t("enrollment.page.preview", "Проверка и Word")
    ConfigureScrollablePage pgPreview, 540

    CreateEmployeePage
    CreateDocsPage
    CreateMonthlyPage
    CreateOneTimePage
    CreateAdvancedPage
    CreateExtrasPage
    CreatePreviewPage
End Sub

Private Sub CreateEmployeePage()
    Set txtEmployeeFIO = AddPageTextBoxT(pgEmployee, "enrollment.field.fio", "ФИО", 12, 12, 520)
    Set txtEmployeeNumber = AddPageTextBoxT(pgEmployee, "enrollment.field.personal_number", "Личный номер", 12, 54, 130)
    Set txtEmployeeTableNumber = AddPageTextBoxT(pgEmployee, "enrollment.field.table_number", "Табельный номер", 160, 54, 120)
    Set txtEmployeeRank = AddPageTextBoxT(pgEmployee, "enrollment.field.rank", "Воинское звание", 298, 54, 120)
    Set txtEmployeeServiceCategory = AddPageTextBoxT(pgEmployee, "enrollment.field.service_category", "Категория службы", 436, 54, 96)
    Set txtEmployeeContractKind = AddPageTextBoxT(pgEmployee, "enrollment.field.contract_kind", "Признак контракта", 12, 96, 130)
    Set txtEmployeeContractBasis = AddPageTextBoxT(pgEmployee, "enrollment.field.contract_basis", "Основание контракта", 160, 96, 372)
    Set txtEmployeeVus = AddPageTextBoxT(pgEmployee, "enrollment.field.vus", "ВУС", 12, 138, 120)
    Set txtEmployeePosition = AddPageTextBoxT(pgEmployee, "enrollment.field.position", "Штатная должность", 150, 138, 382, 34, True)
    Set txtEmployeeSection = AddPageTextBoxT(pgEmployee, "enrollment.field.section", "Раздел персонала", 12, 192, 250, 34, True)
    Set txtEmployeeMilitaryUnit = AddPageTextBoxT(pgEmployee, "enrollment.field.military_unit", "Воинская часть", 280, 192, 252, 34, True)
    Set txtEmployeeTariff = AddPageTextBoxT(pgEmployee, "enrollment.field.tariff", "Тарифный разряд", 12, 246, 120)
    Set txtEmployeePositionSalary = AddPageTextBoxT(pgEmployee, "enrollment.field.position_salary", "Оклад по должности", 150, 246, 140)
    Set txtEmployeeRankSalary = AddPageTextBoxT(pgEmployee, "enrollment.field.rank_salary", "Оклад по званию", 308, 246, 140)
End Sub

Private Sub CreateDocsPage()
    Set txtOrderDraftId = AddPageTextBoxT(pgDocs, "enrollment.field.order_draft_id", "OrderDraftId", 12, 12, 180)
    Set txtOrderDate = AddPageTextBoxT(pgDocs, "enrollment.field.order_date", "Дата приказа", 210, 12, 120)
    Set txtOrderNumber = AddPageTextBoxT(pgDocs, "enrollment.field.order_number", "Номер приказа", 348, 12, 90)
    Set txtOrderIssuer = AddPageTextBoxT(pgDocs, "enrollment.field.order_issuer", "Кем издан приказ", 12, 54, 520)
    Set txtArrivalSource = AddPageTextBoxT(pgDocs, "enrollment.field.arrival_source", "Пункт отбора / источник прибытия", 12, 96, 520, 34, True)
    Set txtPrescriptionNumber = AddPageTextBoxT(pgDocs, "enrollment.field.prescription_number", "Номер предписания", 12, 150, 120)
    Set txtPrescriptionDate = AddPageTextBoxT(pgDocs, "enrollment.field.prescription_date", "Дата предписания", 150, 150, 120)
    Set txtReportNumber = AddPageTextBoxT(pgDocs, "enrollment.field.report_number", "Номер рапорта", 288, 150, 120)
    Set txtReportDate = AddPageTextBoxT(pgDocs, "enrollment.field.report_date", "Дата рапорта", 426, 150, 106)
    Set txtReportInfo = AddPageTextBoxT(pgDocs, "enrollment.field.report_info", "Рапорт / регистрация", 12, 192, 520, 34, True)
    Set txtAssignmentInfo = AddPageTextBoxT(pgDocs, "enrollment.field.assignment_info", "Предписание / основание", 12, 246, 520, 34, True)
    Set txtAcceptDate = AddPageTextBoxT(pgDocs, "enrollment.field.accept_date", "Дата принятия дел и должности", 12, 300, 160)
    Set txtEnrollDate = AddPageTextBoxT(pgDocs, "enrollment.field.enroll_date", "Дата зачисления", 188, 300, 120)
    Set txtDutyStartDate = AddPageTextBoxT(pgDocs, "enrollment.field.duty_start_date", "Дата вступления в обязанности", 324, 300, 160)
    Set txtManualStart = AddPageTextBoxT(pgDocs, "enrollment.field.manual_start", "Ручная дата старта", 12, 342, 160)
    Set txtStandardStart = AddPageTextBoxT(pgDocs, "enrollment.field.standard_start", "Старт стандартных выплат", 188, 342, 160)
    Set txtPreferentialStart = AddPageTextBoxT(pgDocs, "enrollment.field.preferential_start", "Старт льготной выслуги", 364, 342, 168)
    Set txtBasisSection1 = AddPageTextBoxT(pgDocs, "enrollment.field.basis_section1", "Основание для §1", 12, 384, 520, 34, True)
    Set txtBasisSection2 = AddPageTextBoxT(pgDocs, "enrollment.field.basis_section2", "Основание для §2", 12, 438, 520, 34, True)
End Sub

Private Sub CreateMonthlyPage()
    Set chkPreferential = AddPageCheckBoxT(pgMonthly, "enrollment.field.preferential_enabled", "Льготная выслуга", 12, 12)
    Set txtPreferentialCoeff = AddPageTextBoxT(pgMonthly, "enrollment.field.preferential_coeff", "Коэффициент", 180, 8, 70)
    Set chkStdDuty = AddPageCheckBoxT(pgMonthly, "enrollment.field.std_duty", "Надбавка по должности", 12, 54)
    Set txtStdDutyPercent = AddPageTextBoxT(pgMonthly, "common.percent", "%", 210, 50, 60)
    Set chkStdSpecial = AddPageCheckBoxT(pgMonthly, "enrollment.field.std_special", "Особые условия", 290, 54)
    Set txtStdSpecialPercent = AddPageTextBoxT(pgMonthly, "common.percent", "%", 480, 50, 50)
    Set chkStdTariff = AddPageCheckBoxT(pgMonthly, "enrollment.field.std_tariff", "1-4 тариф", 12, 96)
    Set txtStdTariffPercent = AddPageTextBoxT(pgMonthly, "common.percent", "%", 210, 92, 60)
    Set chkStdContract430 = AddPageCheckBoxT(pgMonthly, "enrollment.field.std_contract430", "430 ДСП / контракт", 290, 96)
    Set txtStdContract430Percent = AddPageTextBoxT(pgMonthly, "common.percent", "%", 480, 92, 50)
    Set chkPremium = AddPageCheckBoxT(pgMonthly, "enrollment.field.premium", "Премия", 12, 138)
    Set txtPremiumPercent = AddPageTextBoxT(pgMonthly, "enrollment.field.premium_percent", "Премия %", 80, 134, 60)
    Set txtPremiumStart = AddPageTextBoxT(pgMonthly, "enrollment.field.premium_start", "Начало премии", 160, 134, 120)
    Set txtPremiumEnd = AddPageTextBoxT(pgMonthly, "enrollment.field.premium_end", "Окончание премии", 300, 134, 120)

    Set txtClassParam = AddPageTextBoxT(pgMonthly, "enrollment.field.class_param", "Классность", 12, 188, 120)
    Set chkClass = AddPageCheckBoxT(pgMonthly, "common.enabled_short", "Вкл", 150, 206)
    Set txtClassPercent = AddPageTextBoxT(pgMonthly, "common.percent", "%", 210, 188, 60)
    Set txtFizoParam = AddPageTextBoxT(pgMonthly, "enrollment.field.fizo_param", "ФИЗО", 290, 188, 120)
    Set chkFizo = AddPageCheckBoxT(pgMonthly, "common.enabled_short", "Вкл", 428, 206)
    Set txtFizoPercent = AddPageTextBoxT(pgMonthly, "common.percent", "%", 480, 188, 50)

    Set txtSecrecyParam = AddPageTextBoxT(pgMonthly, "enrollment.field.secrecy_param", "Секретность", 12, 242, 180)
    Set chkSecrecy = AddPageCheckBoxT(pgMonthly, "common.enabled_short", "Вкл", 210, 260)
    Set txtSecrecyPercent = AddPageTextBoxT(pgMonthly, "common.percent", "%", 260, 242, 60)
    Set txtAchievementParam = AddPageTextBoxT(pgMonthly, "enrollment.field.achievement_param", "Особые достижения / медаль", 12, 296, 348, 34, True)
    Set chkAchievement = AddPageCheckBoxT(pgMonthly, "common.enabled_short", "Вкл", 380, 314)
    Set txtAchievementAmount = AddPageTextBoxT(pgMonthly, "enrollment.field.achievement_amount", "% / сумма", 430, 296, 100)
End Sub

Private Sub CreateOneTimePage()
    Set chkLift = AddPageCheckBoxT(pgOneTime, "enrollment.field.lift_enabled", "Подъёмное пособие", 12, 12)
    Set txtLiftAmount = AddPageTextBoxT(pgOneTime, "common.amount", "Сумма", 130, 8, 120)
    Set txtLiftDate = AddPageTextBoxT(pgOneTime, "common.date", "Дата", 270, 8, 120)
    Set chkPerDiem = AddPageCheckBoxT(pgOneTime, "enrollment.field.per_diem_enabled", "Суточные", 12, 54)
    Set txtPerDiemDays = AddPageTextBoxT(pgOneTime, "common.days", "Дни", 130, 50, 60)
    Set txtPerDiemAmount = AddPageTextBoxT(pgOneTime, "common.amount", "Сумма", 210, 50, 120)
    Set txtPerDiemDate = AddPageTextBoxT(pgOneTime, "common.date", "Дата", 350, 50, 120)
    Set chkEdv = AddPageCheckBoxT(pgOneTime, "enrollment.field.edv_enabled", "ЕДВ 400000", 12, 96)
    Set txtEdvAmount = AddPageTextBoxT(pgOneTime, "common.amount", "Сумма", 130, 92, 120)
    Set txtEdvDate = AddPageTextBoxT(pgOneTime, "common.date", "Дата", 270, 92, 120)
    Set txtBirthDate = AddPageTextBoxT(pgOneTime, "enrollment.field.birth_date", "Дата рождения", 12, 146, 120)
    Set txtBirthPlace = AddPageTextBoxT(pgOneTime, "enrollment.field.birth_place", "Место рождения", 150, 146, 320)
    Set txtCitizenship = AddPageTextBoxT(pgOneTime, "enrollment.field.citizenship", "Гражданство", 12, 188, 120)
    Set txtInn = AddPageTextBoxT(pgOneTime, "enrollment.field.inn", "ИНН", 150, 188, 120)
    Set txtSnils = AddPageTextBoxT(pgOneTime, "enrollment.field.snils", "СНИЛС", 290, 188, 180)
    Set txtPassportSeries = AddPageTextBoxT(pgOneTime, "enrollment.field.passport_series", "Серия паспорта", 12, 230, 120)
    Set txtPassportNumber = AddPageTextBoxT(pgOneTime, "enrollment.field.passport_number", "Номер паспорта", 150, 230, 120)
    Set txtPassportIssueDate = AddPageTextBoxT(pgOneTime, "enrollment.field.passport_issue_date", "Дата выдачи", 290, 230, 120)
    Set txtPassportCode = AddPageTextBoxT(pgOneTime, "enrollment.field.passport_code", "Код подразделения", 430, 230, 102)
    Set txtPassportIssuer = AddPageTextBoxT(pgOneTime, "enrollment.field.passport_issuer", "Кем выдан", 550, 146, 250, 50, True)
    Set txtBankAccount = AddPageTextBoxT(pgOneTime, "enrollment.field.bank_account", "Лицевой / банковский счёт", 550, 218, 250)
    Set txtBankName = AddPageTextBoxT(pgOneTime, "enrollment.field.bank_name", "Банк", 550, 260, 250)
    Set txtRequisitesNote = AddPageTextBoxT(pgOneTime, "enrollment.field.requisites_note", "Примечание по реквизитам", 550, 302, 250, 64, True)
End Sub

Private Sub CreateAdvancedPage()
    Set txtPreferentialBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.preferential_basis", "Основание льготной выслуги", 12, 12, 248, 34, True)
    Set txtPremiumBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.premium_basis", "Основание премии", 280, 12, 248, 34, True)
    Set txtStdDutyBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.std_duty_basis", "Основание надбавки по должности", 12, 68, 248, 34, True)
    Set txtStdSpecialBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.std_special_basis", "Основание особых условий", 280, 68, 248, 34, True)
    Set txtStdTariffBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.std_tariff_basis", "Основание тарифной надбавки", 12, 124, 248, 34, True)
    Set txtStdContract430Basis = AddPageTextBoxT(pgAdvanced, "enrollment.field.std_contract430_basis", "Основание 430 ДСП", 280, 124, 248, 34, True)
    Set txtClassBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.class_basis", "Основание классности", 12, 180, 248, 34, True)
    Set txtFizoBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.fizo_basis", "Основание ФИЗО", 280, 180, 248, 34, True)
    Set txtSecrecyBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.secrecy_basis", "Основание секретности", 12, 236, 248, 34, True)
    Set txtAchievementBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.achievement_basis", "Основание особых достижений", 280, 236, 248, 34, True)
    Set txtLiftBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.lift_basis", "Основание подъёмного пособия", 12, 292, 248, 34, True)
    Set txtPerDiemBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.per_diem_basis", "Основание суточных", 280, 292, 248, 34, True)
    Set txtEdvBasis = AddPageTextBoxT(pgAdvanced, "enrollment.field.edv_basis", "Основание ЕДВ", 12, 348, 516, 34, True)
End Sub

Private Sub CreateExtrasPage()
    Dim i As Long
    Dim topPos As Single
    Const MONTHLY_STEP As Single = 126
    Const ONE_TIME_STEP As Single = 120
    Const ONE_TIME_START As Single = 536

    For i = 1 To 4
        topPos = 12 + (i - 1) * MONTHLY_STEP
        Set txtExtraMonthlyName(i) = AddPageTextBox(pgExtras, tf("enrollment.field.extra_monthly_name_short", "Ежемес. #{index}: вид", "{index}", i), 12, topPos, 300)
        txtExtraMonthlyName(i).Tag = CStr(i)
        Set chkExtraMonthly(i) = AddPageCheckBoxT(pgExtras, "common.enabled_short", "Вкл", 324, topPos + 18)
        chkExtraMonthly(i).Width = 44
        Set txtExtraMonthlyParam(i) = AddPageTextBoxT(pgExtras, "enrollment.field.extra_monthly_param", "Параметр", 386, topPos, 130)
        Set txtExtraMonthlyAmount(i) = AddPageTextBoxT(pgExtras, "enrollment.field.extra_monthly_amount", "Размер", 534, topPos, 88)
        Set txtExtraMonthlyStart(i) = AddPageTextBoxT(pgExtras, "enrollment.field.extra_monthly_start", "Дата начала", 640, topPos, 96)
        Set txtExtraMonthlyBasis(i) = AddPageTextBoxT(pgExtras, "enrollment.field.extra_monthly_basis", "Основание", 12, topPos + 52, 724, 42, True)
    Next i

    For i = 1 To 3
        topPos = ONE_TIME_START + (i - 1) * ONE_TIME_STEP
        Set txtExtraOneTimeName(i) = AddPageTextBox(pgExtras, tf("enrollment.field.extra_onetime_name_short", "Разовая #{index}: вид", "{index}", i), 12, topPos, 300)
        txtExtraOneTimeName(i).Tag = CStr(i)
        Set chkExtraOneTime(i) = AddPageCheckBoxT(pgExtras, "common.enabled_short", "Вкл", 324, topPos + 18)
        chkExtraOneTime(i).Width = 44
        Set txtExtraOneTimeAmount(i) = AddPageTextBoxT(pgExtras, "enrollment.field.extra_onetime_amount", "Сумма", 386, topPos, 130)
        Set txtExtraOneTimeDate(i) = AddPageTextBoxT(pgExtras, "enrollment.field.extra_onetime_date", "Дата", 534, topPos, 96)
        Set txtExtraOneTimeBasis(i) = AddPageTextBoxT(pgExtras, "enrollment.field.extra_onetime_basis", "Основание", 12, topPos + 52, 724, 42, True)
    Next i

    ConfigureScrollablePage pgExtras, 920
End Sub

Private Sub CreatePreviewPage()
    Set txtPreviewStatus = AddPageTextBoxT(pgPreview, "enrollment.preview.status", "Статус", 12, 12, 140, 18, False, True)
    Set txtPreviewReady = AddPageTextBoxT(pgPreview, "enrollment.preview.word_ready", "WordReady", 170, 12, 80, 18, False, True)
    Set txtPreviewIssues = AddPageTextBoxT(pgPreview, "enrollment.preview.issues", "Замечания", 12, 54, 520, 90, True, True)
    Set txtPreviewStandard = AddPageTextBoxT(pgPreview, "enrollment.preview.standard", "Стандартные выплаты", 12, 164, 520, 60, True, True)
    Set txtPreviewPersonal = AddPageTextBoxT(pgPreview, "enrollment.preview.personal", "Именные выплаты", 12, 244, 520, 60, True, True)
    Set txtPreviewSection1 = AddPageTextBoxT(pgPreview, "enrollment.preview.section1", "Текст §1", 12, 324, 520, 80, True, True)
    Set txtPreviewSection2 = AddPageTextBoxT(pgPreview, "enrollment.preview.section2", "Текст §2", 12, 424, 520, 60, True, True)
End Sub

Private Sub RemoveDefaultWizardPages()
    Do While mpWizard.Pages.Count > 1
        mpWizard.Pages.Remove mpWizard.Pages.Count - 1
    Loop
End Sub

Private Sub ConfigureScrollablePage(ByVal pageHost As Object, ByVal pageScrollHeight As Long)
    On Error Resume Next
    pageHost.ScrollBars = fmScrollBarsVertical
    pageHost.ScrollHeight = pageScrollHeight
    On Error GoTo 0
End Sub

Private Function AddPageTextBoxT(ByVal pageHost As Object, ByVal localizationKey As String, ByVal fallbackText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal textWidth As Single, Optional ByVal textHeight As Single = 18, Optional ByVal isMultiline As Boolean = False, Optional ByVal isReadOnly As Boolean = False) As Object
    Set AddPageTextBoxT = AddPageTextBox(pageHost, t(localizationKey, fallbackText), leftPos, topPos, textWidth, textHeight, isMultiline, isReadOnly)
End Function

Private Function AddPageCheckBoxT(ByVal pageHost As Object, ByVal localizationKey As String, ByVal fallbackText As String, ByVal leftPos As Single, ByVal topPos As Single) As Object
    Set AddPageCheckBoxT = AddPageCheckBox(pageHost, t(localizationKey, fallbackText), leftPos, topPos)
End Function

Private Function AddPageTextBox(ByVal pageHost As Object, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal textWidth As Single, Optional ByVal textHeight As Single = 18, Optional ByVal isMultiline As Boolean = False, Optional ByVal isReadOnly As Boolean = False) As Object
    Dim lbl As Object
    Dim txt As Object
    Dim safeName As String

    safeName = Replace$(Replace$(Replace$(captionText, " ", "_"), "%", "pct"), "/", "_")

    Set lbl = pageHost.Controls.Add("Forms.Label.1", "lbl_" & safeName & "_" & CStr(pageHost.Controls.Count + 1), True)
    With lbl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = textWidth
        .Height = 14
        .BackStyle = fmBackStyleTransparent
        .Font.Name = "Times New Roman"
        .Font.Size = 9
    End With

    Set txt = pageHost.Controls.Add("Forms.TextBox.1", "txt_" & safeName & "_" & CStr(pageHost.Controls.Count + 1), True)
    With txt
        .Left = leftPos
        .Top = topPos + 15
        .Width = textWidth
        .Height = textHeight
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .MultiLine = isMultiline
        .EnterKeyBehavior = isMultiline
        .WordWrap = isMultiline
        .Locked = isReadOnly
        If isReadOnly Then
            .BackColor = RGB(242, 242, 242)
        Else
            .BackColor = RGB(255, 255, 255)
        End If
    End With

    Set AddPageTextBox = txt
End Function

Private Function AddPageCheckBox(ByVal pageHost As Object, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single) As Object
    Dim chk As Object

    Set chk = pageHost.Controls.Add("Forms.CheckBox.1", "chk_" & CStr(pageHost.Controls.Count + 1), True)
    With chk
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = 150
        .Height = 18
        .Font.Name = "Times New Roman"
        .Font.Size = 10
    End With

    Set AddPageCheckBox = chk
End Function

Private Sub AddSearchResult(ByVal wsStaff As Worksheet, ByVal rowNum As Long, ByVal listRow As Long)
    lstResults.AddItem Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colLichniyNomer_Global).Value))
    lstResults.List(listRow, RESULT_COL_FIO) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colFIO_Global).Value))
    lstResults.List(listRow, RESULT_COL_RANK) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colZvanie_Global).Value))
    lstResults.List(listRow, RESULT_COL_POSITION) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colDolzhnost_Global).Value))
    lstResults.List(listRow, RESULT_COL_SECTION) = Trim$(CStr(wsStaff.Cells(rowNum, mdlHelper.colVoinskayaChast_Global).Value))
End Sub

Private Sub PushFormToBackend()
    mdlEnrollmentWorkflow.SetBackendValue "source_mode", currentSourceMode
    mdlEnrollmentWorkflow.SetBackendValue "fio", txtEmployeeFIO.Value
    mdlEnrollmentWorkflow.SetBackendValue "personal_number", txtEmployeeNumber.Value
    mdlEnrollmentWorkflow.SetBackendValue "table_number", txtEmployeeTableNumber.Value
    mdlEnrollmentWorkflow.SetBackendValue "rank", txtEmployeeRank.Value
    mdlEnrollmentWorkflow.SetBackendValue "service_category", txtEmployeeServiceCategory.Value
    mdlEnrollmentWorkflow.SetBackendValue "contract_kind", txtEmployeeContractKind.Value
    mdlEnrollmentWorkflow.SetBackendValue "contract_basis", txtEmployeeContractBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "vus", txtEmployeeVus.Value
    mdlEnrollmentWorkflow.SetBackendValue "position", txtEmployeePosition.Value
    mdlEnrollmentWorkflow.SetBackendValue "section", txtEmployeeSection.Value
    mdlEnrollmentWorkflow.SetBackendValue "military_unit", txtEmployeeMilitaryUnit.Value
    mdlEnrollmentWorkflow.SetBackendValue "tariff_rank", txtEmployeeTariff.Value
    mdlEnrollmentWorkflow.SetBackendValue "position_salary", txtEmployeePositionSalary.Value
    mdlEnrollmentWorkflow.SetBackendValue "rank_salary", txtEmployeeRankSalary.Value

    mdlEnrollmentWorkflow.SetBackendValue "order_draft_id", txtOrderDraftId.Value
    mdlEnrollmentWorkflow.SetBackendValue "order_date", txtOrderDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "order_number", txtOrderNumber.Value
    mdlEnrollmentWorkflow.SetBackendValue "order_issuer", txtOrderIssuer.Value
    mdlEnrollmentWorkflow.SetBackendValue "arrival_source", txtArrivalSource.Value
    mdlEnrollmentWorkflow.SetBackendValue "prescription_number", txtPrescriptionNumber.Value
    mdlEnrollmentWorkflow.SetBackendValue "prescription_date", txtPrescriptionDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "report_number", txtReportNumber.Value
    mdlEnrollmentWorkflow.SetBackendValue "report_date", txtReportDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "report_info", txtReportInfo.Value
    mdlEnrollmentWorkflow.SetBackendValue "assignment_info", txtAssignmentInfo.Value
    mdlEnrollmentWorkflow.SetBackendValue "accept_date", txtAcceptDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "enroll_date", txtEnrollDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "duty_start_date", txtDutyStartDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "manual_start_date", txtManualStart.Value
    mdlEnrollmentWorkflow.SetBackendValue "standard_start_date", txtStandardStart.Value
    mdlEnrollmentWorkflow.SetBackendValue "preferential_start_date", txtPreferentialStart.Value
    mdlEnrollmentWorkflow.SetBackendValue "basis_section1", txtBasisSection1.Value
    mdlEnrollmentWorkflow.SetBackendValue "basis_section2", txtBasisSection2.Value

    mdlEnrollmentWorkflow.SetBackendValue "preferential_enabled", CheckValue(chkPreferential.Value)
    mdlEnrollmentWorkflow.SetBackendValue "preferential_coeff", txtPreferentialCoeff.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_duty_enabled", CheckValue(chkStdDuty.Value)
    mdlEnrollmentWorkflow.SetBackendValue "std_duty_percent", txtStdDutyPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_special_enabled", CheckValue(chkStdSpecial.Value)
    mdlEnrollmentWorkflow.SetBackendValue "std_special_percent", txtStdSpecialPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_tariff_enabled", CheckValue(chkStdTariff.Value)
    mdlEnrollmentWorkflow.SetBackendValue "std_tariff_percent", txtStdTariffPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_contract430_enabled", CheckValue(chkStdContract430.Value)
    mdlEnrollmentWorkflow.SetBackendValue "std_contract430_percent", txtStdContract430Percent.Value
    mdlEnrollmentWorkflow.SetBackendValue "premium_enabled", CheckValue(chkPremium.Value)
    mdlEnrollmentWorkflow.SetBackendValue "premium_percent", txtPremiumPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "premium_start", txtPremiumStart.Value
    mdlEnrollmentWorkflow.SetBackendValue "premium_end", txtPremiumEnd.Value
    mdlEnrollmentWorkflow.SetBackendValue "premium_basis", txtPremiumBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "class_param", txtClassParam.Value
    mdlEnrollmentWorkflow.SetBackendValue "class_enabled", CheckValue(chkClass.Value)
    mdlEnrollmentWorkflow.SetBackendValue "class_percent", txtClassPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "class_basis", txtClassBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "fizo_param", txtFizoParam.Value
    mdlEnrollmentWorkflow.SetBackendValue "fizo_enabled", CheckValue(chkFizo.Value)
    mdlEnrollmentWorkflow.SetBackendValue "fizo_percent", txtFizoPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "fizo_basis", txtFizoBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "secrecy_param", txtSecrecyParam.Value
    mdlEnrollmentWorkflow.SetBackendValue "secrecy_enabled", CheckValue(chkSecrecy.Value)
    mdlEnrollmentWorkflow.SetBackendValue "secrecy_percent", txtSecrecyPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "secrecy_basis", txtSecrecyBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "achievement_param", txtAchievementParam.Value
    mdlEnrollmentWorkflow.SetBackendValue "achievement_enabled", CheckValue(chkAchievement.Value)
    mdlEnrollmentWorkflow.SetBackendValue "achievement_amount", txtAchievementAmount.Value
    mdlEnrollmentWorkflow.SetBackendValue "achievement_basis", txtAchievementBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "preferential_basis", txtPreferentialBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_duty_basis", txtStdDutyBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_special_basis", txtStdSpecialBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_tariff_basis", txtStdTariffBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_contract430_basis", txtStdContract430Basis.Value
    PushExtraPaymentsToBackend

    mdlEnrollmentWorkflow.SetBackendValue "lift_enabled", CheckValue(chkLift.Value)
    mdlEnrollmentWorkflow.SetBackendValue "lift_amount", txtLiftAmount.Value
    mdlEnrollmentWorkflow.SetBackendValue "lift_date", txtLiftDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "lift_basis", txtLiftBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "per_diem_enabled", CheckValue(chkPerDiem.Value)
    mdlEnrollmentWorkflow.SetBackendValue "per_diem_days", txtPerDiemDays.Value
    mdlEnrollmentWorkflow.SetBackendValue "per_diem_amount", txtPerDiemAmount.Value
    mdlEnrollmentWorkflow.SetBackendValue "per_diem_date", txtPerDiemDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "per_diem_basis", txtPerDiemBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "edv_enabled", CheckValue(chkEdv.Value)
    mdlEnrollmentWorkflow.SetBackendValue "edv_amount", txtEdvAmount.Value
    mdlEnrollmentWorkflow.SetBackendValue "edv_date", txtEdvDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "edv_basis", txtEdvBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "birth_date", txtBirthDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "birth_place", txtBirthPlace.Value
    mdlEnrollmentWorkflow.SetBackendValue "citizenship", txtCitizenship.Value
    mdlEnrollmentWorkflow.SetBackendValue "inn", txtInn.Value
    mdlEnrollmentWorkflow.SetBackendValue "snils", txtSnils.Value
    mdlEnrollmentWorkflow.SetBackendValue "passport_series", txtPassportSeries.Value
    mdlEnrollmentWorkflow.SetBackendValue "passport_number", txtPassportNumber.Value
    mdlEnrollmentWorkflow.SetBackendValue "passport_issuer", txtPassportIssuer.Value
    mdlEnrollmentWorkflow.SetBackendValue "passport_issue_date", txtPassportIssueDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "passport_code", txtPassportCode.Value
    mdlEnrollmentWorkflow.SetBackendValue "bank_account", txtBankAccount.Value
    mdlEnrollmentWorkflow.SetBackendValue "bank_name", txtBankName.Value
    mdlEnrollmentWorkflow.SetBackendValue "requisites_note", txtRequisitesNote.Value
End Sub

Private Function CheckValue(ByVal rawValue As Variant) As String
    If CBool(rawValue) Then
        CheckValue = "YES"
    Else
        CheckValue = "NO"
    End If
End Function

Private Sub PushExtraPaymentsToBackend()
    Dim i As Long

    For i = 1 To 4
        mdlEnrollmentWorkflow.SetBackendValue ExtraMonthlyKey(i, "name"), txtExtraMonthlyName(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraMonthlyKey(i, "param"), txtExtraMonthlyParam(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraMonthlyKey(i, "amount"), txtExtraMonthlyAmount(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraMonthlyKey(i, "start"), txtExtraMonthlyStart(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraMonthlyKey(i, "basis"), txtExtraMonthlyBasis(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraMonthlyKey(i, "enabled"), CheckValue(chkExtraMonthly(i).Value)
    Next i

    For i = 1 To 3
        mdlEnrollmentWorkflow.SetBackendValue ExtraOneTimeKey(i, "name"), txtExtraOneTimeName(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraOneTimeKey(i, "amount"), txtExtraOneTimeAmount(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraOneTimeKey(i, "date"), txtExtraOneTimeDate(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraOneTimeKey(i, "basis"), txtExtraOneTimeBasis(i).Value
        mdlEnrollmentWorkflow.SetBackendValue ExtraOneTimeKey(i, "enabled"), CheckValue(chkExtraOneTime(i).Value)
    Next i
End Sub

Private Sub ReloadExtraPaymentsFromBackend()
    Dim i As Long

    For i = 1 To 4
        txtExtraMonthlyName(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraMonthlyKey(i, "name")))
        txtExtraMonthlyParam(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraMonthlyKey(i, "param")))
        txtExtraMonthlyAmount(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraMonthlyKey(i, "amount")))
        txtExtraMonthlyStart(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraMonthlyKey(i, "start")))
        txtExtraMonthlyBasis(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraMonthlyKey(i, "basis")))
        chkExtraMonthly(i).Value = BackendYesNo(ExtraMonthlyKey(i, "enabled"))
    Next i

    For i = 1 To 3
        txtExtraOneTimeName(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraOneTimeKey(i, "name")))
        txtExtraOneTimeAmount(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraOneTimeKey(i, "amount")))
        txtExtraOneTimeDate(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraOneTimeKey(i, "date")))
        txtExtraOneTimeBasis(i).Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue(ExtraOneTimeKey(i, "basis")))
        chkExtraOneTime(i).Value = BackendYesNo(ExtraOneTimeKey(i, "enabled"))
    Next i
End Sub

Private Function ExtraMonthlyKey(ByVal index As Long, ByVal fieldName As String) As String
    ExtraMonthlyKey = "extra_monthly" & CStr(index) & "_" & fieldName
End Function

Private Function ExtraOneTimeKey(ByVal index As Long, ByVal fieldName As String) As String
    ExtraOneTimeKey = "extra_one_time" & CStr(index) & "_" & fieldName
End Function

Public Sub ReloadFromBackend()
    currentSourceMode = SafeText(mdlEnrollmentWorkflow.GetBackendValue("source_mode"))
    If currentSourceMode = "" Then currentSourceMode = "manual"

    txtEmployeeFIO.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("fio"))
    txtEmployeeNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("personal_number"))
    txtEmployeeTableNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("table_number"))
    txtEmployeeRank.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("rank"))
    txtEmployeeServiceCategory.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("service_category"))
    txtEmployeeContractKind.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("contract_kind"))
    txtEmployeeContractBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("contract_basis"))
    txtEmployeeVus.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("vus"))
    txtEmployeePosition.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("position"))
    txtEmployeeSection.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("section"))
    txtEmployeeMilitaryUnit.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("military_unit"))
    txtEmployeeTariff.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("tariff_rank"))
    txtEmployeePositionSalary.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("position_salary"))
    txtEmployeeRankSalary.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("rank_salary"))

    txtOrderDraftId.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_draft_id"))
    txtOrderDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_date"))
    txtOrderNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_number"))
    txtOrderIssuer.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_issuer"))
    txtArrivalSource.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("arrival_source"))
    txtPrescriptionNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("prescription_number"))
    txtPrescriptionDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("prescription_date"))
    txtReportNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("report_number"))
    txtReportDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("report_date"))
    txtReportInfo.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("report_info"))
    txtAssignmentInfo.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("assignment_info"))
    txtAcceptDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("accept_date"))
    txtEnrollDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("enroll_date"))
    txtDutyStartDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("duty_start_date"))
    txtManualStart.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("manual_start_date"))
    txtStandardStart.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("standard_start_date"))
    txtPreferentialStart.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preferential_start_date"))
    txtBasisSection1.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("basis_section1"))
    txtBasisSection2.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("basis_section2"))

    chkPreferential.Value = BackendYesNo("preferential_enabled")
    txtPreferentialCoeff.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preferential_coeff"))
    chkStdDuty.Value = BackendYesNo("std_duty_enabled")
    txtStdDutyPercent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_duty_percent"))
    chkStdSpecial.Value = BackendYesNo("std_special_enabled")
    txtStdSpecialPercent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_special_percent"))
    chkStdTariff.Value = BackendYesNo("std_tariff_enabled")
    txtStdTariffPercent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_tariff_percent"))
    chkStdContract430.Value = BackendYesNo("std_contract430_enabled")
    txtStdContract430Percent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_contract430_percent"))
    chkPremium.Value = BackendYesNo("premium_enabled")
    txtPremiumPercent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("premium_percent"))
    txtPremiumStart.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("premium_start"))
    txtPremiumEnd.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("premium_end"))
    txtPremiumBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("premium_basis"))
    txtClassParam.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("class_param"))
    chkClass.Value = BackendYesNo("class_enabled")
    txtClassPercent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("class_percent"))
    txtClassBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("class_basis"))
    txtFizoParam.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("fizo_param"))
    chkFizo.Value = BackendYesNo("fizo_enabled")
    txtFizoPercent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("fizo_percent"))
    txtFizoBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("fizo_basis"))
    txtSecrecyParam.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("secrecy_param"))
    chkSecrecy.Value = BackendYesNo("secrecy_enabled")
    txtSecrecyPercent.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("secrecy_percent"))
    txtSecrecyBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("secrecy_basis"))
    txtAchievementParam.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("achievement_param"))
    chkAchievement.Value = BackendYesNo("achievement_enabled")
    txtAchievementAmount.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("achievement_amount"))
    txtAchievementBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("achievement_basis"))
    txtPreferentialBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preferential_basis"))
    txtStdDutyBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_duty_basis"))
    txtStdSpecialBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_special_basis"))
    txtStdTariffBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_tariff_basis"))
    txtStdContract430Basis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("std_contract430_basis"))
    ReloadExtraPaymentsFromBackend

    chkLift.Value = BackendYesNo("lift_enabled")
    txtLiftAmount.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("lift_amount"))
    txtLiftDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("lift_date"))
    txtLiftBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("lift_basis"))
    chkPerDiem.Value = BackendYesNo("per_diem_enabled")
    txtPerDiemDays.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("per_diem_days"))
    txtPerDiemAmount.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("per_diem_amount"))
    txtPerDiemDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("per_diem_date"))
    txtPerDiemBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("per_diem_basis"))
    chkEdv.Value = BackendYesNo("edv_enabled")
    txtEdvAmount.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("edv_amount"))
    txtEdvDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("edv_date"))
    txtEdvBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("edv_basis"))
    txtBirthDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("birth_date"))
    txtBirthPlace.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("birth_place"))
    txtCitizenship.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("citizenship"))
    txtInn.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("inn"))
    txtSnils.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("snils"))
    txtPassportSeries.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("passport_series"))
    txtPassportNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("passport_number"))
    txtPassportIssuer.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("passport_issuer"))
    txtPassportIssueDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("passport_issue_date"))
    txtPassportCode.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("passport_code"))
    txtBankAccount.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("bank_account"))
    txtBankName.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("bank_name"))
    txtRequisitesNote.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("requisites_note"))

    txtPreviewStatus.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_status"))
    txtPreviewReady.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_word_ready"))
    txtPreviewIssues.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_issues"))
    txtPreviewStandard.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_standard"))
    txtPreviewPersonal.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_personal"))
    txtPreviewSection1.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_section1"))
    txtPreviewSection2.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_section2"))
End Sub

Private Function BackendYesNo(ByVal fieldKey As String) As Boolean
    BackendYesNo = (UCase$(SafeText(mdlEnrollmentWorkflow.GetBackendValue(fieldKey))) = "YES")
End Function

Private Function SafeText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    SafeText = Trim$(CStr(rawValue))
End Function
