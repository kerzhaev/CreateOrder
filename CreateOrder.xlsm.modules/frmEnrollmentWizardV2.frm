VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEnrollmentWizardV2 
   Caption         =   "Мастер зачисления V2 - layout"
   ClientHeight    =   15450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17355
   OleObjectBlob   =   "frmEnrollmentWizardV2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEnrollmentWizardV2"
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

' ----- Layout constants for the monthly payments page (CreateMonthlyPage) -----
' Geometry is collected here so the layout can be tuned without reading the body of CreateMonthlyPage.
Private Const FRA727_LEFT As Single = 12
Private Const FRA727_TOP As Single = 52
Private Const FRA727_WIDTH As Single = 744
Private Const FRA727_HEIGHT As Single = 168

Private Const FRA430_LEFT As Single = 12
Private Const FRA430_TOP As Single = 232
Private Const FRA430_WIDTH As Single = 744
Private Const FRA430_HEIGHT As Single = 174

' Standard control sizes used across the monthly page.
Private Const CTRL_PERCENT_WIDTH As Single = 28
Private Const CTRL_PARAM_WIDTH As Single = 96
Private Const CTRL_PARAM_ACHIEVEMENT_WIDTH As Single = 120
Private Const CTRL_DATE_WIDTH As Single = 92
Private Const CTRL_DOC_WIDTH As Single = 160
Private Const CTRL_PREMIUM_WIDTH As Single = 82
Private Const CTRL_PREMIUM_START_WIDTH As Single = 82
Private Const CTRL_PREMIUM_END_WIDTH As Single = 82
Private Const CTRL_HEIGHT_DEFAULT As Single = 18
Private Const CTRL_LBL_TARIFF_WIDTH As Single = 246

' Row baselines inside fraOrder727.
' Each row carries a checkbox (top1) and adjacent percent textbox (top2). The
' textbox is placed 4 points above the checkbox so their visible centres align.
Private Const FRA727_ROW1_CHK_TOP As Single = 18       ' std_duty / std_special checkboxes
Private Const FRA727_ROW1_TXT_TOP As Single = 14       ' std_duty% / std_special% textboxes
Private Const FRA727_ROW2_CHK_TOP As Single = 80       ' secrecy / class checkboxes ("Вкл")
Private Const FRA727_ROW2_TXT_TOP As Single = 62       ' secrecy/class param combo and % textbox
Private Const FRA727_ROW3_CHK_TOP As Single = 126      ' premium checkbox (rendered slightly lower than textboxes)
Private Const FRA727_ROW3_TXT_TOP As Single = 110      ' premium% / premium dates
Private Const FRA727_COL_LEFT As Single = 18
Private Const FRA727_COL_LEFT_PERCENT As Single = 178
Private Const FRA727_COL_RIGHT As Single = 386
Private Const FRA727_COL_RIGHT_PERCENT As Single = 514
Private Const FRA727_COL_RIGHT_CHK As Single = 530
Private Const FRA727_COL_RIGHT_CLASS_PERCENT As Single = 578

' Row baselines inside fraOrder430.
Private Const FRA430_ROW1_CHK_TOP As Single = 18       ' contract430 checkbox / contract430% textbox
Private Const FRA430_ROW1_TXT_TOP As Single = 14       ' contract430% textbox (rendered above the checkbox)
Private Const FRA430_ROW2_LBL_TOP As Single = 54       ' 1-4 tariff label / tariff% textbox
Private Const FRA430_ROW2_TXT_TOP As Single = 50       ' tariff% textbox
Private Const FRA430_ROW3_CHK_TOP As Single = 102      ' fizo / achievement checkboxes ("Вкл")
Private Const FRA430_ROW3_TXT_TOP As Single = 84       ' fizo/achievement param combo and % textbox
Private Const FRA430_ROW4_TOP As Single = 126          ' medal award date / order row
Private Const FRA430_COL_LEFT As Single = 18
Private Const FRA430_COL_LEFT_PERCENT As Single = 168
Private Const FRA430_COL_LEFT_TARIFF_PERCENT As Single = 274
Private Const FRA430_COL_LEFT_FIZO_PERCENT As Single = 210
Private Const FRA430_COL_RIGHT As Single = 386
Private Const FRA430_COL_RIGHT_CHK As Single = 568
Private Const FRA430_COL_RIGHT_AMOUNT As Single = 616
Private Const FRA430_COL_RIGHT_AMOUNT_WIDTH As Single = 36
Private Const FRA430_COL_DATE As Single = 386
Private Const FRA430_COL_DOC As Single = 494

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
Private txtEmployeeContractBasis As Object
Private txtEmployeeVus As Object
Private txtEmployeePosition As Object
Private txtEmployeeSection As Object
Private txtEmployeeTariff As Object
Private txtEmployeePositionSalary As Object
Private txtEmployeeRankSalary As Object
Private lblReferenceAmountHint As Object

Private txtOrderDate As Object
Private txtOrderDraftId As Object
Private txtOrderNumber As Object
Private txtOrderIssuer As Object
Private chkArrivalDetails As Object
Private chkReportDetails As Object
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
Private txtAchievementAwardDate As Object
Private txtAchievementDocumentReference As Object
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
Private chkPersonalDetails As Object
Private chkBankDetails As Object
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
Private txtBankBik As Object
Private txtRequisitesNote As Object

Private txtPremiumBasis As Object
Private txtLiftBasis As Object
Private txtPerDiemBasis As Object
Private txtEdvBasis As Object
Private txtClassBasis As Object
Private txtFizoBasis As Object
Private txtSecrecyBasis As Object
Private txtAchievementBasis As Object
Private txtStdContract430Basis As Object
Private lblTariffAllowanceState As Object

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
Private txtPreviewOutputPath As Object

Private WithEvents cboEmployeeRankDynamic As MSForms.ComboBox
Private WithEvents cboEmployeeTariffDynamic As MSForms.ComboBox
Private WithEvents cboClassDynamic As MSForms.ComboBox
Private WithEvents cboSecrecyDynamic As MSForms.ComboBox
Private WithEvents cboFizoDynamic As MSForms.ComboBox
Private WithEvents cboAchievementDynamic As MSForms.ComboBox
Private WithEvents cboBankDynamic As MSForms.ComboBox

Private currentSourceMode As String
Private Const PREVIEW_PAGE_INDEX As Long = 6

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    ' Infrastructure is prepared by the public workflow command before Show.
    mdlHelper.EnsureStaffColumnsInitialized
    BindDesignerControls
    ConfigureSearchArea
    ConfigureWindow
    ConfigureButtons
    ApplyDesignerLocalization
    PopulateOperatorReferenceLists
    currentSourceMode = "manual"
    ReloadFromBackend
    UpdatePaymentBasisHighlights
    lblStatus.Caption = t("enrollment.form.status.ready_to_pick", "Выберите сотрудника из листа 'Штат' или заполните карточку вручную. После выбора проверьте страницы мастера.")
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "frmEnrollmentWizardV2.UserForm_Initialize", Err.Description
End Sub

Private Sub BindDesignerControls()
    Set pgEmployee = mpWizard.Pages("pgEmployee")
    Set pgDocs = mpWizard.Pages("pgDocs")
    Set pgMonthly = mpWizard.Pages("pgMonthly")
    Set pgOneTime = mpWizard.Pages("pgOneTime")
    Set pgAdvanced = mpWizard.Pages("pgAdvanced")
    Set pgExtras = mpWizard.Pages("pgExtras")
    Set pgPreview = mpWizard.Pages("pgPreview")
    Set txtEmployeeFIO = FindDesignerControl("txt_____2")
    Set txtEmployeeNumber = FindDesignerControl("txt______________4")
    Set txtEmployeeTableNumber = FindDesignerControl("txt_________________6")
    Set txtEmployeeRank = FindDesignerControl("cbo_8")
    Set txtEmployeeServiceCategory = FindDesignerControl("cbo_10")
    Set txtEmployeeContractBasis = FindDesignerControl("txt______________________________12")
    Set txtEmployeeVus = FindDesignerControl("cbo_14")
    Set txtEmployeePosition = FindDesignerControl("cbo_16")
    Set txtEmployeeSection = FindDesignerControl("cbo_18")
    Set txtEmployeeTariff = FindDesignerControl("cbo_20")
    Set txtEmployeePositionSalary = FindDesignerControl("txt____________________22")
    Set txtEmployeeRankSalary = FindDesignerControl("txt_________________24")
    Set lblReferenceAmountHint = FindDesignerControl("lbl_hint_25")
    Set txtOrderDate = FindDesignerControl("doc_txt______________4")
    Set txtOrderDraftId = FindDesignerControl("txt_OrderDraftId_2")
    Set txtOrderNumber = FindDesignerControl("txt_______________6")
    Set txtOrderIssuer = FindDesignerControl("txt__________________8")
    Set chkArrivalDetails = FindDesignerControl("chk_1")
    Set chkReportDetails = FindDesignerControl("doc_chk_1")
    Set txtArrivalSource = FindDesignerControl("txt__________________________________3")
    Set txtPrescriptionNumber = FindDesignerControl("txt___________________5")
    Set txtPrescriptionDate = FindDesignerControl("txt__________________7")
    Set txtReportNumber = FindDesignerControl("txt_______________3")
    Set txtReportDate = FindDesignerControl("txt______________5")
    Set txtReportInfo = FindDesignerControl("txt______________________7")
    Set txtAssignmentInfo = FindDesignerControl("txt_________________________9")
    Set txtAcceptDate = FindDesignerControl("txt_______________________________2")
    Set txtEnrollDate = FindDesignerControl("txt_________________4")
    Set txtDutyStartDate = FindDesignerControl("txt_______________________________6")
    Set txtManualStart = FindDesignerControl("txt____________________8")
    Set txtStandardStart = FindDesignerControl("txt__________________________10")
    Set txtPreferentialStart = FindDesignerControl("txt________________________12")
    Set txtBasisSection1 = FindDesignerControl("txt________________1_2")
    Set txtBasisSection2 = FindDesignerControl("txt________________2_4")
    Set chkStdDuty = FindDesignerControl("mon_chk_1_2")
    Set txtStdDutyPercent = FindDesignerControl("txt_pct_3")
    Set chkStdSpecial = FindDesignerControl("chk_4")
    Set txtStdSpecialPercent = FindDesignerControl("txt_pct_6")
    Set chkStdTariff = FindDesignerControl("mon_chk_4")
    Set txtStdTariffPercent = FindDesignerControl("txt_pct_7")
    Set chkStdContract430 = FindDesignerControl("mon_chk_1_3")
    Set txtStdContract430Percent = FindDesignerControl("mon_txt_pct_3")
    Set chkPremium = FindDesignerControl("chk_17")
    Set txtPremiumPercent = FindDesignerControl("txt________pct_19")
    Set txtPremiumStart = FindDesignerControl("txt_______________21")
    Set txtPremiumEnd = FindDesignerControl("txt_______________________23")
    Set txtClassParam = FindDesignerControl("cbo_13")
    Set chkClass = FindDesignerControl("chk_14")
    Set txtClassPercent = FindDesignerControl("txt_pct_16")
    Set txtFizoParam = FindDesignerControl("cbo_9")
    Set chkFizo = FindDesignerControl("chk_10")
    Set txtFizoPercent = FindDesignerControl("txt_pct_12")
    Set txtSecrecyParam = FindDesignerControl("mon_cbo_8")
    Set chkSecrecy = FindDesignerControl("chk_9")
    Set txtSecrecyPercent = FindDesignerControl("txt_pct_11")
    Set txtAchievementParam = FindDesignerControl("mon_cbo_14")
    Set chkAchievement = FindDesignerControl("chk_15")
    Set txtAchievementAmount = FindDesignerControl("txt_pct_________17")
    Set txtAchievementAwardDate = FindDesignerControl("txt______19")
    Set txtAchievementDocumentReference = FindDesignerControl("mon_txt_______________21")
    Set chkPreferential = FindDesignerControl("mon_chk_1")
    Set txtPreferentialCoeff = FindDesignerControl("txt_____________3")
    Set chkLift = FindDesignerControl("chk_2")
    Set txtLiftAmount = FindDesignerControl("txt_______4")
    Set txtLiftDate = FindDesignerControl("txt______6")
    Set chkPerDiem = FindDesignerControl("chk_7")
    Set txtPerDiemDays = FindDesignerControl("txt_____9")
    Set txtPerDiemAmount = FindDesignerControl("txt_______11")
    Set txtPerDiemDate = FindDesignerControl("txt______13")
    Set chkEdv = FindDesignerControl("one_chk_14")
    Set txtEdvAmount = FindDesignerControl("txt_______16")
    Set txtEdvDate = FindDesignerControl("txt______18")
    Set chkPersonalDetails = FindDesignerControl("chk_19")
    Set chkBankDetails = FindDesignerControl("chk_20")
    Set txtBirthDate = FindDesignerControl("txt_______________22")
    Set txtBirthPlace = FindDesignerControl("txt________________24")
    Set txtCitizenship = FindDesignerControl("txt_____________26")
    Set txtInn = FindDesignerControl("txt_____28")
    Set txtSnils = FindDesignerControl("txt_______30")
    Set txtPassportSeries = FindDesignerControl("txt________________32")
    Set txtPassportNumber = FindDesignerControl("txt________________34")
    Set txtPassportIssuer = FindDesignerControl("txt___________40")
    Set txtPassportIssueDate = FindDesignerControl("txt_____________36")
    Set txtPassportCode = FindDesignerControl("txt___________________38")
    Set txtBankAccount = FindDesignerControl("txt___________________________46")
    Set txtBankName = FindDesignerControl("cbo_42")
    Set txtBankBik = FindDesignerControl("txt______________________44")
    Set txtRequisitesNote = FindDesignerControl("txt__________________________48")
    Set txtPremiumBasis = FindDesignerControl("txt__________________6")
    Set txtLiftBasis = FindDesignerControl("txt______________________________2")
    Set txtPerDiemBasis = FindDesignerControl("txt____________________4")
    Set txtEdvBasis = FindDesignerControl("adv_txt_______________6")
    Set txtClassBasis = FindDesignerControl("txt______________________4")
    Set txtFizoBasis = FindDesignerControl("txt________________4")
    Set txtSecrecyBasis = FindDesignerControl("txt_______________________2")
    Set txtAchievementBasis = FindDesignerControl("txt_____________________________6")
    Set txtStdContract430Basis = FindDesignerControl("txt___________430_____2")
    Set lblTariffAllowanceState = FindDesignerControl("lbl_hint_5")
    Set txtExtraMonthlyName(1) = FindDesignerControl("cbo_3")
    Set txtExtraMonthlyName(2) = FindDesignerControl("ext_cbo_14")
    Set txtExtraMonthlyName(3) = FindDesignerControl("cbo_25")
    Set txtExtraMonthlyName(4) = FindDesignerControl("cbo_36")
    Set txtExtraMonthlyParam(1) = FindDesignerControl("txt__________6")
    Set txtExtraMonthlyParam(2) = FindDesignerControl("txt__________17")
    Set txtExtraMonthlyParam(3) = FindDesignerControl("txt__________28")
    Set txtExtraMonthlyParam(4) = FindDesignerControl("txt__________39")
    Set txtExtraMonthlyAmount(1) = FindDesignerControl("txt________8")
    Set txtExtraMonthlyAmount(2) = FindDesignerControl("txt________19")
    Set txtExtraMonthlyAmount(3) = FindDesignerControl("txt________30")
    Set txtExtraMonthlyAmount(4) = FindDesignerControl("txt________41")
    Set txtExtraMonthlyStart(1) = FindDesignerControl("txt_____________10")
    Set txtExtraMonthlyStart(2) = FindDesignerControl("txt_____________21")
    Set txtExtraMonthlyStart(3) = FindDesignerControl("txt_____________32")
    Set txtExtraMonthlyStart(4) = FindDesignerControl("txt_____________43")
    Set txtExtraMonthlyBasis(1) = FindDesignerControl("txt___________12")
    Set txtExtraMonthlyBasis(2) = FindDesignerControl("txt___________23")
    Set txtExtraMonthlyBasis(3) = FindDesignerControl("txt___________34")
    Set txtExtraMonthlyBasis(4) = FindDesignerControl("txt___________45")
    Set chkExtraMonthly(1) = FindDesignerControl("ext_chk_4")
    Set chkExtraMonthly(2) = FindDesignerControl("ext_chk_15")
    Set chkExtraMonthly(3) = FindDesignerControl("chk_26")
    Set chkExtraMonthly(4) = FindDesignerControl("chk_37")
    Set txtExtraOneTimeName(1) = FindDesignerControl("cbo_48")
    Set txtExtraOneTimeName(2) = FindDesignerControl("cbo_57")
    Set txtExtraOneTimeName(3) = FindDesignerControl("cbo_66")
    Set txtExtraOneTimeAmount(1) = FindDesignerControl("txt_______51")
    Set txtExtraOneTimeAmount(2) = FindDesignerControl("txt_______60")
    Set txtExtraOneTimeAmount(3) = FindDesignerControl("txt_______69")
    Set txtExtraOneTimeDate(1) = FindDesignerControl("txt______53")
    Set txtExtraOneTimeDate(2) = FindDesignerControl("txt______62")
    Set txtExtraOneTimeDate(3) = FindDesignerControl("txt______71")
    Set txtExtraOneTimeBasis(1) = FindDesignerControl("txt___________55")
    Set txtExtraOneTimeBasis(2) = FindDesignerControl("txt___________64")
    Set txtExtraOneTimeBasis(3) = FindDesignerControl("txt___________73")
    Set chkExtraOneTime(1) = FindDesignerControl("chk_49")
    Set chkExtraOneTime(2) = FindDesignerControl("chk_58")
    Set chkExtraOneTime(3) = FindDesignerControl("chk_67")
    Set txtPreviewStatus = FindDesignerControl("txt________3")
    Set txtPreviewReady = FindDesignerControl("txt_Word_______5")
    Set txtPreviewIssues = FindDesignerControl("txt_______________9")
    Set txtPreviewStandard = FindDesignerControl("txt________________727_11")
    Set txtPreviewPersonal = FindDesignerControl("txt________________430____13")
    Set txtPreviewSection1 = FindDesignerControl("txt_______________________1_15")
    Set txtPreviewSection2 = FindDesignerControl("txt_______________________2_17")
    Set txtPreviewOutputPath = FindDesignerControl("txt______Word________________7")
    Set cboEmployeeRankDynamic = FindDesignerControl("cbo_8")
    Set cboEmployeeTariffDynamic = FindDesignerControl("cbo_20")
    Set cboClassDynamic = FindDesignerControl("cbo_13")
    Set cboSecrecyDynamic = FindDesignerControl("mon_cbo_8")
    Set cboFizoDynamic = FindDesignerControl("cbo_9")
    Set cboAchievementDynamic = FindDesignerControl("mon_cbo_14")
    Set cboBankDynamic = FindDesignerControl("cbo_42")
End Sub

Private Function FindDesignerControl(ByVal controlName As String) As Object
    Set FindDesignerControl = FindDesignerControlInContainer(Me, controlName)
    If FindDesignerControl Is Nothing Then Err.Raise 5, "frmEnrollmentWizardV2.FindDesignerControl", "Не найден design-time контрол: " & controlName
End Function

Private Function FindDesignerControlInContainer(ByVal containerHost As Object, ByVal controlName As String) As Object
    Dim controlItem As Object
    Dim pageItem As Object
    Dim foundControl As Object

    On Error Resume Next
    Set foundControl = containerHost.Controls.Item(controlName)
    On Error GoTo 0
    If Not foundControl Is Nothing Then
        Set FindDesignerControlInContainer = foundControl
        Exit Function
    End If

    For Each controlItem In containerHost.Controls
        If TypeName(controlItem) = "MultiPage" Then
            For Each pageItem In controlItem.Pages
                Set foundControl = FindDesignerControlInContainer(pageItem, controlName)
                If Not foundControl Is Nothing Then
                    Set FindDesignerControlInContainer = foundControl
                    Exit Function
                End If
            Next pageItem
        ElseIf TypeName(controlItem) = "Frame" Then
            Set foundControl = FindDesignerControlInContainer(controlItem, controlName)
            If Not foundControl Is Nothing Then
                Set FindDesignerControlInContainer = foundControl
                Exit Function
            End If
        End If
    Next controlItem
End Function
Private Sub ApplyDesignerLocalization()
    Me.Caption = t("enrollment.form.title", "Мастер зачисления")
    pgEmployee.Caption = t("enrollment.page.employee", "Военнослужащий")
    pgDocs.Caption = t("enrollment.page.docs", "Документы и даты")
    pgMonthly.Caption = t("enrollment.page.monthly", "Ежемесячные выплаты")
    pgOneTime.Caption = t("enrollment.page.onetime", "Разовые выплаты и реквизиты")
    pgAdvanced.Caption = t("enrollment.page.advanced", "Основания выплат")
    pgExtras.Caption = t("enrollment.page.extras", "Иные выплаты")
    pgPreview.Caption = t("enrollment.page.preview", "Проверка и текст приказа")
    SetDesignerCaption "lbl_____1", "enrollment.field.fio", "ФИО", 0
    SetDesignerCaption "lbl______________3", "enrollment.field.personal_number", "Личный номер", 0
    SetDesignerCaption "lbl_________________5", "enrollment.field.table_number", "Табельный номер", 0
    SetDesignerCaption "lbl_cbo_7", "enrollment.field.rank", "Воинское звание", 0
    SetDesignerCaption "lbl_cbo_9", "enrollment.field.service_category", "Категория службы", 0
    SetDesignerCaption "lbl_cbo_13", "enrollment.field.vus", "ВУС", 0
    SetDesignerCaption "lbl_cbo_15", "enrollment.field.position", "Штатная должность", 0
    SetDesignerCaption "lbl_cbo_17", "enrollment.field.section", "Раздел персонала", 0
    SetDesignerCaption "lbl_cbo_19", "enrollment.field.tariff", "Тарифный разряд", 0
    SetDesignerCaption "lbl____________________21", "enrollment.field.position_salary", "Оклад по должности (из справочника)", 0
    SetDesignerCaption "lbl_________________23", "enrollment.field.rank_salary", "Оклад по званию (из справочника)", 0
    SetDesignerCaption "lbl_OrderDraftId_1", "enrollment.field.order_draft_id", "Проект", 0
    SetDesignerCaption "doc_lbl______________3", "enrollment.field.order_date", "Дата приказа", 0
    SetDesignerCaption "lbl_______________5", "enrollment.field.order_number", "Номер", 0
    SetDesignerCaption "lbl__________________7", "enrollment.field.order_issuer", "Кем издан", 0
    SetDesignerCaption "chk_1", "enrollment.field.arrival_details_enabled", "Внести сведения о прибытии", 0
    SetDesignerCaption "lbl__________________________________2", "enrollment.field.arrival_source", "Источник прибытия", 0
    SetDesignerCaption "lbl___________________4", "enrollment.field.prescription_number", "Предписание №", 0
    SetDesignerCaption "lbl__________________6", "enrollment.field.prescription_date", "Дата", 0
    SetDesignerCaption "lbl_________________________8", "enrollment.field.assignment_info", "Основание / примечание", 0
    SetDesignerCaption "doc_chk_1", "enrollment.field.report_details_enabled", "Внести сведения о рапорте", 0
    SetDesignerCaption "lbl_______________2", "enrollment.field.report_number", "Рапорт №", 0
    SetDesignerCaption "lbl______________4", "enrollment.field.report_date", "Дата", 0
    SetDesignerCaption "lbl______________________6", "enrollment.field.report_info", "Рапорт / регистрация", 0
    SetDesignerCaption "lbl_______________________________1", "enrollment.field.accept_date", "Принял дела", 0
    SetDesignerCaption "lbl_________________3", "enrollment.field.enroll_date", "Зачислен", 0
    SetDesignerCaption "lbl_______________________________5", "enrollment.field.duty_start_date", "Приступил", 0
    SetDesignerCaption "lbl____________________7", "enrollment.field.manual_start", "Ручной старт", 0
    SetDesignerCaption "lbl__________________________9", "enrollment.field.standard_start", "Старт выплат", 0
    SetDesignerCaption "lbl________________________11", "enrollment.field.preferential_start", "Старт выслуги", 0
    SetDesignerCaption "lbl________________1_1", "enrollment.field.basis_section1", "Пункт 1", 0
    SetDesignerCaption "lbl________________2_3", "enrollment.field.basis_section2", "Пункт 2", 0
    SetDesignerCaption "mon_chk_1", "enrollment.field.preferential_enabled", "Льготная выслуга", 0
    SetDesignerCaption "lbl_____________2", "enrollment.field.preferential_coeff", "Коэффициент", 0
    SetDesignerCaption "mon_chk_1_2", "enrollment.field.std_duty", "Надбавка по должности", 0
    SetDesignerCaption "lbl_pct_2", "common.percent", "%", 0
    SetDesignerCaption "chk_4", "enrollment.field.std_special", "Особые условия", 0
    SetDesignerCaption "lbl_pct_5", "common.percent", "%", 0
    SetDesignerCaption "mon_lbl_cbo_7", "enrollment.field.secrecy_param", "Секретность", 0
    SetDesignerCaption "chk_9", "common.enabled_short", "Вкл", 0
    SetDesignerCaption "lbl_pct_10", "common.percent", "%", 0
    SetDesignerCaption "lbl_cbo_12", "enrollment.field.class_param", "Классность", 0
    SetDesignerCaption "chk_14", "common.enabled_short", "Вкл", 0
    SetDesignerCaption "lbl_pct_15", "common.percent", "%", 0
    SetDesignerCaption "chk_17", "enrollment.field.premium", "Премия", 0
    SetDesignerCaption "lbl________pct_18", "enrollment.field.premium_percent", "%", 0
    SetDesignerCaption "lbl_______________20", "enrollment.field.premium_start", "Начало", 0
    SetDesignerCaption "lbl_______________________22", "enrollment.field.premium_end", "Окончание", 0
    SetDesignerCaption "mon_chk_1_3", "enrollment.field.std_contract430", "Контракт / 430дсп", 0
    SetDesignerCaption "mon_lbl_pct_2", "common.percent", "%", 0
    SetDesignerCaption "mon_chk_4", "enrollment.field.std_tariff", "", 0
    SetDesignerCaption "lbl_pct_6", "common.percent", "%", 0
    SetDesignerCaption "lbl_cbo_8", "enrollment.field.fizo_param", "ФИЗО", 0
    SetDesignerCaption "chk_10", "common.enabled_short", "Вкл", 0
    SetDesignerCaption "lbl_pct_11", "common.percent", "%", 0
    SetDesignerCaption "mon_lbl_cbo_13", "enrollment.field.achievement_param", "Особые достижения / медаль", 0
    SetDesignerCaption "chk_15", "common.enabled_short", "Вкл", 0
    SetDesignerCaption "lbl_pct_________16", "enrollment.field.achievement_amount", "% / сумма", 0
    SetDesignerCaption "lbl______18", "common.date", "Дата приказа", 0
    SetDesignerCaption "mon_lbl_______________20", "enrollment.field.order_number", "Номер приказа", 0
    SetDesignerCaption "chk_2", "enrollment.field.lift_enabled", "Подъёмное пособие", 0
    SetDesignerCaption "lbl_______3", "common.amount", "Сумма", 0
    SetDesignerCaption "lbl______5", "common.date", "Дата", 0
    SetDesignerCaption "chk_7", "enrollment.field.per_diem_enabled", "Суточные", 0
    SetDesignerCaption "lbl_____8", "common.days", "Дни", 0
    SetDesignerCaption "lbl_______10", "common.amount", "Сумма", 0
    SetDesignerCaption "lbl______12", "common.date", "Дата", 0
    SetDesignerCaption "one_chk_14", "enrollment.field.edv_enabled", "ЕДВ 400000", 0
    SetDesignerCaption "lbl_______15", "common.amount", "Сумма", 0
    SetDesignerCaption "lbl______17", "common.date", "Дата", 0
    SetDesignerCaption "chk_19", "enrollment.field.personal_details_enabled", "Внести персональные данные", 0
    SetDesignerCaption "chk_20", "enrollment.field.bank_details_enabled", "Внести банковские реквизиты", 0
    SetDesignerCaption "lbl_______________21", "enrollment.field.birth_date", "Дата рождения", 0
    SetDesignerCaption "lbl________________23", "enrollment.field.birth_place", "Место рождения", 0
    SetDesignerCaption "lbl_____________25", "enrollment.field.citizenship", "Гражданство", 0
    SetDesignerCaption "lbl_____27", "enrollment.field.inn", "ИНН", 0
    SetDesignerCaption "lbl_______29", "enrollment.field.snils", "СНИЛС", 0
    SetDesignerCaption "lbl________________31", "enrollment.field.passport_series", "Серия паспорта", 0
    SetDesignerCaption "lbl________________33", "enrollment.field.passport_number", "Номер паспорта", 0
    SetDesignerCaption "lbl_____________35", "enrollment.field.passport_issue_date", "Дата выдачи", 0
    SetDesignerCaption "lbl___________________37", "enrollment.field.passport_code", "Код подразделения", 0
    SetDesignerCaption "lbl___________39", "enrollment.field.passport_issuer", "Кем выдан", 0
    SetDesignerCaption "lbl_cbo_41", "enrollment.field.bank_name", "Банк (из справочника)", 0
    SetDesignerCaption "lbl___________________________45", "enrollment.field.bank_account", "Лицевой / банковский счёт", 0
    SetDesignerCaption "lbl__________________________47", "enrollment.field.requisites_note", "Примечание по реквизитам", 0
    SetDesignerCaption "lbl_______________________1", "enrollment.field.secrecy_basis", "Секретность", 0
    SetDesignerCaption "lbl______________________3", "enrollment.field.class_basis", "Классность", 0
    SetDesignerCaption "lbl__________________5", "enrollment.field.premium_basis", "Премия", 0
    SetDesignerCaption "lbl___________430_____1", "enrollment.field.std_contract430_basis", "Контракт / 430дсп", 0
    SetDesignerCaption "lbl________________3", "enrollment.field.fizo_basis", "ФИЗО", 0
    SetDesignerCaption "lbl_____________________________5", "enrollment.field.achievement_basis", "Особые достижения / медаль", 0
    SetDesignerCaption "lbl______________________________1", "enrollment.field.lift_basis", "Подъёмное пособие", 0
    SetDesignerCaption "lbl____________________3", "enrollment.field.per_diem_basis", "Суточные", 0
    SetDesignerCaption "adv_lbl_______________5", "enrollment.field.edv_basis", "ЕДВ", 0
    SetDesignerCaption "lbl_cbo_2", "enrollment.field.extra_monthly_name_short", "Вид", 1
    SetDesignerCaption "ext_lbl_cbo_13", "enrollment.field.extra_monthly_name_short", "Вид", 2
    SetDesignerCaption "lbl_cbo_24", "enrollment.field.extra_monthly_name_short", "Вид", 3
    SetDesignerCaption "lbl_cbo_35", "enrollment.field.extra_monthly_name_short", "Вид", 4
    SetDesignerCaption "ext_chk_4", "common.enabled_short", "Вкл", 1
    SetDesignerCaption "ext_chk_15", "common.enabled_short", "Вкл", 2
    SetDesignerCaption "chk_26", "common.enabled_short", "Вкл", 3
    SetDesignerCaption "chk_37", "common.enabled_short", "Вкл", 4
    SetDesignerCaption "lbl__________5", "enrollment.field.extra_monthly_param", "Приказ / пункт", 1
    SetDesignerCaption "lbl__________16", "enrollment.field.extra_monthly_param", "Приказ / пункт", 2
    SetDesignerCaption "lbl__________27", "enrollment.field.extra_monthly_param", "Приказ / пункт", 3
    SetDesignerCaption "lbl__________38", "enrollment.field.extra_monthly_param", "Приказ / пункт", 4
    SetDesignerCaption "lbl________7", "enrollment.field.extra_monthly_amount", "Размер", 1
    SetDesignerCaption "lbl________18", "enrollment.field.extra_monthly_amount", "Размер", 2
    SetDesignerCaption "lbl________29", "enrollment.field.extra_monthly_amount", "Размер", 3
    SetDesignerCaption "lbl________40", "enrollment.field.extra_monthly_amount", "Размер", 4
    SetDesignerCaption "lbl_____________9", "enrollment.field.extra_monthly_start", "Дата", 1
    SetDesignerCaption "lbl_____________20", "enrollment.field.extra_monthly_start", "Дата", 2
    SetDesignerCaption "lbl_____________31", "enrollment.field.extra_monthly_start", "Дата", 3
    SetDesignerCaption "lbl_____________42", "enrollment.field.extra_monthly_start", "Дата", 4
    SetDesignerCaption "lbl___________11", "enrollment.field.extra_monthly_basis", "Основание", 1
    SetDesignerCaption "lbl___________22", "enrollment.field.extra_monthly_basis", "Основание", 2
    SetDesignerCaption "lbl___________33", "enrollment.field.extra_monthly_basis", "Основание", 3
    SetDesignerCaption "lbl___________44", "enrollment.field.extra_monthly_basis", "Основание", 4
    SetDesignerCaption "lbl_cbo_47", "enrollment.field.extra_onetime_name_short", "Вид", 1
    SetDesignerCaption "lbl_cbo_56", "enrollment.field.extra_onetime_name_short", "Вид", 2
    SetDesignerCaption "lbl_cbo_65", "enrollment.field.extra_onetime_name_short", "Вид", 3
    SetDesignerCaption "chk_49", "common.enabled_short", "Вкл", 1
    SetDesignerCaption "chk_58", "common.enabled_short", "Вкл", 2
    SetDesignerCaption "chk_67", "common.enabled_short", "Вкл", 3
    SetDesignerCaption "lbl_______50", "enrollment.field.extra_onetime_amount", "Сумма", 1
    SetDesignerCaption "lbl_______59", "enrollment.field.extra_onetime_amount", "Сумма", 2
    SetDesignerCaption "lbl_______68", "enrollment.field.extra_onetime_amount", "Сумма", 3
    SetDesignerCaption "lbl______52", "enrollment.field.extra_onetime_date", "Дата", 1
    SetDesignerCaption "lbl______61", "enrollment.field.extra_onetime_date", "Дата", 2
    SetDesignerCaption "lbl______70", "enrollment.field.extra_onetime_date", "Дата", 3
    SetDesignerCaption "lbl___________54", "enrollment.field.extra_onetime_basis", "Приказ / основание", 1
    SetDesignerCaption "lbl___________63", "enrollment.field.extra_onetime_basis", "Приказ / основание", 2
    SetDesignerCaption "lbl___________72", "enrollment.field.extra_onetime_basis", "Приказ / основание", 3
End Sub

Private Sub SetDesignerCaption(ByVal controlName As String, ByVal localizationKey As String, ByVal fallbackText As String, Optional ByVal captionIndex As Long = 0)
    Dim targetControl As Object
    Dim resolvedCaption As String
    Set targetControl = FindDesignerControl(controlName)
    resolvedCaption = t(localizationKey, fallbackText)
    If captionIndex > 0 Then resolvedCaption = Replace$(resolvedCaption, "{index}", CStr(captionIndex))
    targetControl.Caption = resolvedCaption
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
        Err.Raise vbObjectError + 1810, "frmEnrollmentWizardV2.LoadEmployeeFromStaffNumber", t("form.select_employee.message.staff_columns_error", "Не удалось определить обязательные столбцы листа 'Штат'.")
    End If

    staffRow = ResolveStaffRowByAnyNumber(wsStaff, employeeNumber)
    If staffRow < 2 Then
        Err.Raise vbObjectError + 1811, "frmEnrollmentWizardV2.LoadEmployeeFromStaffNumber", tf("enrollment.form.error.employee_not_found", "Сотрудник с номером {number} не найден на листе 'Штат'.", "{number}", employeeNumber)
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
    txtEmployeeRank.Value = mdlEnrollmentWorkflow.GetEnrollmentReferenceDisplayNameOrCode("RANK", Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colZvanie_Global).Value)))
    txtEmployeePosition.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colDolzhnost_Global).Value))
    txtEmployeeSection.Value = Trim$(CStr(wsStaff.Cells(staffRow, mdlHelper.colVoinskayaChast_Global).Value))
    txtEmployeeServiceCategory.Value = StaffDictValue(staffData, mdlHelper.Ru(1043, 1088, 1091, 1087, 1087, 1072, 32, 1089, 1086, 1090, 1088, 1091, 1076, 1085, 1080, 1082, 1086, 1074))
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
        Trim$(CStr(txtEmployeeSection.Value)) & "|" & _
        Trim$(CStr(txtEmployeeTableNumber.Value)) & "|" & _
        Trim$(CStr(txtEmployeeServiceCategory.Value)) & "|" & _
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

    PerformExportPackage False
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
        FindPageLabelCaption(pgExtras, 12, 24) & "|" & _
        FindPageLabelCaption(pgExtras, 12, 248)
End Function

Private Function FindPageLabelCaption(ByVal pageHost As Object, ByVal leftPos As Single, ByVal topPos As Single) As String
    Dim controlItem As Object

    For Each controlItem In pageHost.Controls
        If TypeName(controlItem) = "Label" Then
            If CLng(controlItem.Left) = CLng(leftPos) And CLng(controlItem.Top) = CLng(topPos) Then
                FindPageLabelCaption = CStr(controlItem.Caption)
                Exit Function
            End If
        End If
    Next controlItem
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
    Dim personnelEventID As String

    On Error GoTo ExportBlocked

    lblStatus.Caption = t("enrollment.form.status.export_preparing", "Подготавливается Word-приказ. Пожалуйста, подождите.")
    Me.Repaint
    DoEvents

    PushFormToBackend
    targetRow = mdlEnrollmentWorkflow.SaveEnrollmentFormToSheet(False)
    personnelEventID = CStr(mdlEnrollmentWorkflow.GetBackendValue("personnel_event_id"))
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
    mdlEnrollmentEventLink.RegisterEnrollmentOrderForEvent personnelEventID, outputPath, CStr(mdlEnrollmentWorkflow.GetBackendValue("order_number"))
    lblStatus.Caption = tf("enrollment.form.status.exported", "Пакет приказа сформирован: {scope}.", "{scope}", exportScope) & " " & outputPath
    txtPreviewOutputPath.Value = outputPath

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
    ConfigureInlineSearchUi
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
    txtSearch.Visible = True
    txtSearch.ControlTipText = t("enrollment.form.search.tip", "Введите ФИО, личный или табельный номер.")
    btnLoadFromInlineSearchDynamic.Caption = t("enrollment.form.button.load_from_search", "Загрузить из поиска")
    btnLoadFromInlineSearchDynamic.Visible = True
    btnLoadFromInlineSearchDynamic.Enabled = False
    lstResults.Visible = True
End Sub
Private Sub ConfigureButtons()
    btnSelect.Caption = t("enrollment.form.button.pick_from_staff", "Выбрать сотрудника из штата")
    btnCheckDynamic.Caption = t("enrollment.form.button.check", "Проверить и показать")
    btnSaveCardDynamic.Caption = t("enrollment.form.button.save", "Сохранить")
    btnExportPackageDynamic.Caption = t("enrollment.form.button.export", "Экспортировать Word")
    btnSaveContinueDynamic.Caption = t("enrollment.form.button.save_continue_package", "Следующий в пакете")
    btnClose.Caption = t("common.close", "Закрыть")
End Sub
Private Sub PopulateOperatorReferenceLists()
    PopulateComboBox txtEmployeeRank, "RANK"
    PopulateComboBox txtEmployeeServiceCategory, "SERVICE_CATEGORY"
    PopulateComboBox txtEmployeeTariff, "TARIFF_RANK"
    PopulateComboBox txtClassParam, "CLASS"
    PopulateComboBox txtSecrecyParam, "SECRECY"
    PopulateComboBox txtFizoParam, "FIZO"
    PopulateComboBox txtAchievementParam, "ACHIEVEMENT"
    PopulateComboBox txtBankName, "BANK"
    Set cboEmployeeRankDynamic = txtEmployeeRank
    Set cboEmployeeTariffDynamic = txtEmployeeTariff
    Set cboClassDynamic = txtClassParam
    Set cboSecrecyDynamic = txtSecrecyParam
    Set cboFizoDynamic = txtFizoParam
    Set cboAchievementDynamic = txtAchievementParam
    Set cboBankDynamic = txtBankName
    PopulateExtraPaymentTypeLists
End Sub


Private Sub PopulateExtraPaymentTypeLists()
    Dim i As Long

    For i = 1 To 4
        txtExtraMonthlyName(i).Clear
        txtExtraMonthlyName(i).AddItem "Ежемесячная надбавка"
        txtExtraMonthlyName(i).AddItem "Ежемесячная доплата"
        txtExtraMonthlyName(i).AddItem "Иная ежемесячная выплата"
    Next i
    For i = 1 To 3
        txtExtraOneTimeName(i).Clear
        txtExtraOneTimeName(i).AddItem "Разовая выплата"
        txtExtraOneTimeName(i).AddItem "Материальная помощь"
        txtExtraOneTimeName(i).AddItem "Иная разовая выплата"
    Next i
End Sub
Private Sub PopulateComboBox(ByVal comboBox As Object, ByVal referenceType As String)
    Dim values As Collection
    Dim itemValue As Variant
    Set values = mdlEnrollmentWorkflow.GetEnrollmentReferenceValues(referenceType)
    comboBox.Clear
    For Each itemValue In values
        comboBox.AddItem CStr(itemValue)
    Next itemValue
End Sub

Private Sub UpdateReferenceSalaryPreview()
    Dim amountValue As String
    Dim missingValues As String

    amountValue = mdlEnrollmentWorkflow.GetRankReferenceAmount(SafeText(txtEmployeeRank.Value))
    If amountValue <> "" Then
        txtEmployeeRankSalary.Value = amountValue
    ElseIf SafeText(txtEmployeeRank.Value) <> "" Then
        txtEmployeeRankSalary.Value = ""
        missingValues = t("enrollment.hint.rank_amount_missing", "оклад по званию")
    End If
    amountValue = mdlEnrollmentWorkflow.GetTariffRankReferenceAmount(SafeText(txtEmployeeTariff.Value))
    If amountValue <> "" Then
        txtEmployeePositionSalary.Value = amountValue
    ElseIf SafeText(txtEmployeeTariff.Value) <> "" Then
        txtEmployeePositionSalary.Value = ""
        If missingValues <> "" Then missingValues = missingValues & "; "
        missingValues = missingValues & t("enrollment.hint.tariff_amount_missing", "оклад по должности")
    End If
    UpdateAutomaticTariffAllowance
    amountValue = mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("CLASS", SafeText(txtClassParam.Value))
    If amountValue <> "" Then txtClassPercent.Value = amountValue
    amountValue = mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("SECRECY", SafeText(txtSecrecyParam.Value))
    If amountValue <> "" Then txtSecrecyPercent.Value = amountValue
    amountValue = mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("FIZO", SafeText(txtFizoParam.Value))
    If amountValue <> "" Then txtFizoPercent.Value = amountValue
    amountValue = mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("ACHIEVEMENT", SafeText(txtAchievementParam.Value))
    If amountValue <> "" Then txtAchievementAmount.Value = amountValue
    txtBankBik.Value = mdlEnrollmentWorkflow.GetEnrollmentReferenceAmount("BANK", SafeText(txtBankName.Value))
    If missingValues = "" Then
        lblReferenceAmountHint.Caption = ""
    Else
        lblReferenceAmountHint.Caption = t("enrollment.hint.reference_amount_missing", "Для выбранного значения не указан {amount} в справочнике EnrollmentReferenceData (столбец Amount).")
        lblReferenceAmountHint.Caption = Replace$(lblReferenceAmountHint.Caption, "{amount}", missingValues)
    End If
End Sub

Private Sub UpdateAutomaticTariffAllowance()
    Dim tariffCode As String
    Dim isAutomaticAllowance As Boolean

    tariffCode = mdlEnrollmentWorkflow.GetEnrollmentReferenceCodeOrDisplay("TARIFF_RANK", SafeText(txtEmployeeTariff.Value))
    If IsNumeric(tariffCode) Then isAutomaticAllowance = CLng(tariffCode) >= 1 And CLng(tariffCode) <= 4
    chkStdTariff.Value = isAutomaticAllowance
    If isAutomaticAllowance Then
        lblTariffAllowanceState.Caption = "1–4 тариф: надбавка назначается автоматически"
    Else
        lblTariffAllowanceState.Caption = "1–4 тариф: для выбранного разряда не применяется"
    End If
End Sub

Private Sub cboEmployeeRankDynamic_Change()
    UpdateReferenceSalaryPreview
End Sub

Private Sub cboEmployeeTariffDynamic_Change()
    UpdateReferenceSalaryPreview
End Sub

Private Sub cboClassDynamic_Change()
    UpdateReferenceSalaryPreview
    UpdatePaymentBasisHighlights
End Sub

Private Sub cboFizoDynamic_Change()
    UpdateReferenceSalaryPreview
    UpdatePaymentBasisHighlights
End Sub

Private Sub cboBankDynamic_Change()
    UpdateReferenceSalaryPreview
End Sub

Private Sub cboSecrecyDynamic_Change()
    UpdateReferenceSalaryPreview
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt__________________6_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt______________________4_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt________________4_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt_______________________2_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt_____________________________6_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt___________430_____2_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt______________________________2_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub txt____________________4_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub adv_txt_______________6_Change()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_17_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_14_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_10_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_9_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_15_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub mon_chk_1_3_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_2_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_7_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub one_chk_14_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub ext_chk_4_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub ext_chk_15_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_26_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_37_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_49_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_58_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub chk_67_Click()
    UpdatePaymentBasisHighlights
End Sub

Private Sub UpdatePaymentBasisHighlights()
    ApplyPaymentBasisHighlight chkPremium, txtPremiumBasis
    ApplyPaymentBasisHighlight chkClass, txtClassBasis
    ApplyPaymentBasisHighlight chkFizo, txtFizoBasis
    ApplyPaymentBasisHighlight chkSecrecy, txtSecrecyBasis
    ApplyPaymentBasisHighlight chkAchievement, txtAchievementBasis
    ApplyPaymentBasisHighlight chkStdContract430, txtStdContract430Basis
    ApplyPaymentBasisHighlight chkLift, txtLiftBasis
    ApplyPaymentBasisHighlight chkPerDiem, txtPerDiemBasis
    ApplyPaymentBasisHighlight chkEdv, txtEdvBasis
    UpdateExtraPaymentBasisHighlights
End Sub

Private Sub UpdateExtraPaymentBasisHighlights()
    Dim i As Long
    For i = 1 To 4
        ApplyPaymentBasisHighlight chkExtraMonthly(i), txtExtraMonthlyBasis(i)
    Next i
    For i = 1 To 3
        ApplyPaymentBasisHighlight chkExtraOneTime(i), txtExtraOneTimeBasis(i)
    Next i
End Sub

Private Sub ApplyPaymentBasisHighlight(ByVal checkBox As Object, ByVal basisTextBox As Object)
    Dim missing As Boolean
    If checkBox Is Nothing Or basisTextBox Is Nothing Then Exit Sub
    missing = CBool(checkBox.Value) And Len(Trim$(CStr(basisTextBox.Value))) = 0
    If missing Then
        basisTextBox.BackColor = RGB(255, 230, 230)
        basisTextBox.ControlTipText = t("enrollment.issue.payment_field_missing", "Заполните обязательное основание выплаты.")
    Else
        basisTextBox.BackColor = RGB(255, 255, 255)
        basisTextBox.ControlTipText = vbNullString
    End If
End Sub
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
    mdlEnrollmentWorkflow.SetBackendValue "rank", mdlEnrollmentWorkflow.GetEnrollmentReferenceCodeOrDisplay("RANK", txtEmployeeRank.Value)
    mdlEnrollmentWorkflow.SetBackendValue "service_category", txtEmployeeServiceCategory.Value
    mdlEnrollmentWorkflow.SetBackendValue "contract_kind", txtEmployeeServiceCategory.Value
    mdlEnrollmentWorkflow.SetBackendValue "contract_basis", txtEmployeeContractBasis.Value
    mdlEnrollmentWorkflow.SetBackendValue "vus", txtEmployeeVus.Value
    mdlEnrollmentWorkflow.SetBackendValue "position", txtEmployeePosition.Value
    mdlEnrollmentWorkflow.SetBackendValue "section", txtEmployeeSection.Value
    mdlEnrollmentWorkflow.SetBackendValue "military_unit", txtEmployeeSection.Value
    mdlEnrollmentWorkflow.SetBackendValue "tariff_rank", txtEmployeeTariff.Value
    mdlEnrollmentWorkflow.SetBackendValue "position_salary", txtEmployeePositionSalary.Value
    mdlEnrollmentWorkflow.SetBackendValue "rank_salary", txtEmployeeRankSalary.Value

    mdlEnrollmentWorkflow.SetBackendValue "order_draft_id", txtOrderDraftId.Value
    mdlEnrollmentWorkflow.SetBackendValue "order_date", txtOrderDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "order_number", txtOrderNumber.Value
    mdlEnrollmentWorkflow.SetBackendValue "order_issuer", txtOrderIssuer.Value
    mdlEnrollmentWorkflow.SetBackendValue "arrival_details_enabled", CheckValue(chkArrivalDetails.Value)
    If chkArrivalDetails.Value Then
        mdlEnrollmentWorkflow.SetBackendValue "arrival_source", txtArrivalSource.Value
        mdlEnrollmentWorkflow.SetBackendValue "prescription_number", txtPrescriptionNumber.Value
        mdlEnrollmentWorkflow.SetBackendValue "prescription_date", txtPrescriptionDate.Value
        mdlEnrollmentWorkflow.SetBackendValue "assignment_info", txtAssignmentInfo.Value
    Else
        ClearBackendValues Array("arrival_source", "prescription_number", "prescription_date", "assignment_info")
    End If
    mdlEnrollmentWorkflow.SetBackendValue "report_details_enabled", CheckValue(chkReportDetails.Value)
    If chkReportDetails.Value Then
        mdlEnrollmentWorkflow.SetBackendValue "report_number", txtReportNumber.Value
        mdlEnrollmentWorkflow.SetBackendValue "report_date", txtReportDate.Value
        mdlEnrollmentWorkflow.SetBackendValue "report_info", txtReportInfo.Value
    Else
        ClearBackendValues Array("report_number", "report_date", "report_info")
    End If
    mdlEnrollmentWorkflow.SetBackendValue "accept_date", txtAcceptDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "enroll_date", txtEnrollDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "duty_start_date", txtDutyStartDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "manual_start_date", txtManualStart.Value
    mdlEnrollmentWorkflow.SetBackendValue "standard_start_date", txtStandardStart.Value
    mdlEnrollmentWorkflow.SetBackendValue "preferential_start_date", txtPreferentialStart.Value
    mdlEnrollmentWorkflow.SetBackendValue "basis_section1", txtBasisSection1.Value
    mdlEnrollmentWorkflow.SetBackendValue "basis_section2", txtBasisSection2.Value
    ClearBackendValues Array("preferential_basis", "std_duty_basis", "std_special_basis", "std_tariff_basis")

    mdlEnrollmentWorkflow.SetBackendValue "preferential_enabled", CheckValue(chkPreferential.Value)
    mdlEnrollmentWorkflow.SetBackendValue "preferential_coeff", txtPreferentialCoeff.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_duty_enabled", CheckValue(chkStdDuty.Value)
    mdlEnrollmentWorkflow.SetBackendValue "std_duty_percent", txtStdDutyPercent.Value
    mdlEnrollmentWorkflow.SetBackendValue "std_special_enabled", CheckValue(chkStdSpecial.Value)
    mdlEnrollmentWorkflow.SetBackendValue "std_special_percent", txtStdSpecialPercent.Value
    UpdateAutomaticTariffAllowance
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
    mdlEnrollmentWorkflow.SetBackendValue "achievement_award_date", txtAchievementAwardDate.Value
    mdlEnrollmentWorkflow.SetBackendValue "achievement_document_reference", txtAchievementDocumentReference.Value
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
    mdlEnrollmentWorkflow.SetBackendValue "personal_details_enabled", CheckValue(chkPersonalDetails.Value)
    If chkPersonalDetails.Value Then
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
    Else
        ClearBackendValues Array("birth_date", "birth_place", "citizenship", "inn", "snils", "passport_series", "passport_number", "passport_issuer", "passport_issue_date", "passport_code")
    End If
    mdlEnrollmentWorkflow.SetBackendValue "bank_details_enabled", CheckValue(chkBankDetails.Value)
    If chkBankDetails.Value Then
        mdlEnrollmentWorkflow.SetBackendValue "bank_account", txtBankAccount.Value
        mdlEnrollmentWorkflow.SetBackendValue "bank_name", txtBankName.Value
        mdlEnrollmentWorkflow.SetBackendValue "bank_bik", txtBankBik.Value
        mdlEnrollmentWorkflow.SetBackendValue "requisites_note", txtRequisitesNote.Value
    Else
        ClearBackendValues Array("bank_account", "bank_name", "bank_bik", "requisites_note")
    End If
End Sub

Private Sub ClearBackendValues(ByVal fieldKeys As Variant)
    Dim item As Variant

    For Each item In fieldKeys
        mdlEnrollmentWorkflow.SetBackendValue CStr(item), ""
    Next item
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
    txtEmployeeRank.Value = mdlEnrollmentWorkflow.GetEnrollmentReferenceDisplayNameOrCode("RANK", SafeText(mdlEnrollmentWorkflow.GetBackendValue("rank")))
    txtEmployeeServiceCategory.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("service_category"))
    txtEmployeeContractBasis.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("contract_basis"))
    txtEmployeeVus.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("vus"))
    txtEmployeePosition.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("position"))
    txtEmployeeSection.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("section"))
    txtEmployeeTariff.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("tariff_rank"))
    txtEmployeePositionSalary.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("position_salary"))
    txtEmployeeRankSalary.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("rank_salary"))
    UpdateReferenceSalaryPreview

    txtOrderDraftId.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_draft_id"))
    txtOrderDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_date"))
    txtOrderNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_number"))
    txtOrderIssuer.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("order_issuer"))
    chkArrivalDetails.Value = BackendYesNo("arrival_details_enabled")
    txtArrivalSource.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("arrival_source"))
    txtPrescriptionNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("prescription_number"))
    txtPrescriptionDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("prescription_date"))
    chkReportDetails.Value = BackendYesNo("report_details_enabled")
    txtReportNumber.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("report_number"))
    txtReportDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("report_date"))
    txtReportInfo.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("report_info"))
    txtAssignmentInfo.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("assignment_info"))
    chkArrivalDetails.Value = BackendYesNo("arrival_details_enabled") Or SafeText(txtArrivalSource.Value) <> "" Or _
        SafeText(txtPrescriptionNumber.Value) <> "" Or SafeText(txtPrescriptionDate.Value) <> "" Or SafeText(txtAssignmentInfo.Value) <> ""
    chkReportDetails.Value = BackendYesNo("report_details_enabled") Or SafeText(txtReportNumber.Value) <> "" Or _
        SafeText(txtReportDate.Value) <> "" Or SafeText(txtReportInfo.Value) <> ""
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
    txtAchievementAwardDate.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("achievement_award_date"))
    txtAchievementDocumentReference.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("achievement_document_reference"))
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
    chkPersonalDetails.Value = BackendYesNo("personal_details_enabled") Or SafeText(txtBirthDate.Value) <> "" Or _
        SafeText(txtBirthPlace.Value) <> "" Or SafeText(txtCitizenship.Value) <> "" Or SafeText(txtInn.Value) <> "" Or _
        SafeText(txtSnils.Value) <> "" Or SafeText(txtPassportSeries.Value) <> "" Or SafeText(txtPassportNumber.Value) <> "" Or _
        SafeText(txtPassportIssuer.Value) <> "" Or SafeText(txtPassportIssueDate.Value) <> "" Or SafeText(txtPassportCode.Value) <> ""
    txtBankAccount.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("bank_account"))
    txtBankName.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("bank_name"))
    txtBankBik.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("bank_bik"))
    chkBankDetails.Value = BackendYesNo("bank_details_enabled") Or SafeText(txtBankAccount.Value) <> "" Or _
        SafeText(txtBankName.Value) <> "" Or SafeText(txtBankBik.Value) <> ""
    txtRequisitesNote.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("requisites_note"))

    txtPreviewStatus.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_status"))
    txtPreviewReady.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_word_ready"))
    txtPreviewIssues.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_issues"))
    txtPreviewStandard.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_standard"))
    txtPreviewPersonal.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_personal"))
    txtPreviewSection1.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_section1"))
    txtPreviewSection2.Value = SafeText(mdlEnrollmentWorkflow.GetBackendValue("preview_section2"))
    UpdatePaymentBasisHighlights
End Sub

Private Function BackendYesNo(ByVal fieldKey As String) As Boolean
    BackendYesNo = (UCase$(SafeText(mdlEnrollmentWorkflow.GetBackendValue(fieldKey))) = "YES")
End Function

Private Function SafeText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    SafeText = Trim$(CStr(rawValue))
End Function
