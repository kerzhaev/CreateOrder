VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPersonnelActionWizardV2 
   Caption         =   "Кадровое действие V2"
   ClientHeight    =   12180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15555
   OleObjectBlob   =   "frmPersonnelActionWizardV2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPersonnelActionWizardV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEBUG_LOGGING As Boolean = True

Private mSavedEventID As String
Private mSelectionMode As Boolean
Private mLoadedSignature As String
Private mSearchMatches As Collection
Private mPreviewDraft As Object
Private mPreviewResult As Object
Private mPreviewSignature As String
Private mPreviewState As String
Private WithEvents mSearchText As MSForms.TextBox
Attribute mSearchText.VB_VarHelpID = -1
Private WithEvents mMenuEnrollment As MSForms.CommandButton
Attribute mMenuEnrollment.VB_VarHelpID = -1
Private WithEvents mMenuTransfer As MSForms.CommandButton
Attribute mMenuTransfer.VB_VarHelpID = -1
Private WithEvents mMenuExclusion As MSForms.CommandButton
Attribute mMenuExclusion.VB_VarHelpID = -1
Private WithEvents mMenuHistory As MSForms.CommandButton
Attribute mMenuHistory.VB_VarHelpID = -1
Private WithEvents mMenuClose As MSForms.CommandButton
Attribute mMenuClose.VB_VarHelpID = -1
Private WithEvents mPreviewConfirm As MSForms.CommandButton
Attribute mPreviewConfirm.VB_VarHelpID = -1
Private WithEvents mPreviewCancel As MSForms.CommandButton
Attribute mPreviewCancel.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    mSelectionMode = mdlPersonnelEvents.ConsumePersonnelActionMenuRequest()
    LogDebug "initialize", "selection_mode=" & CStr(mSelectionMode)
    BindDesignerControls
    ConfigureWizard
    mPreviewState = "EDITING"
    If Not mSelectionMode Then
        LoadValues
        mLoadedSignature = CurrentSignature
    End If
    LogDebug "initialize-complete", "selection_mode=" & CStr(mSelectionMode)
    Exit Sub

Failed:
    failureNumber = Err.number
    failureDescription = Err.description
    LogDebug "initialize-error", "number=" & CStr(failureNumber) & "; description=" & failureDescription
    Err.Raise failureNumber, "frmPersonnelActionWizardV2.UserForm_Initialize", failureDescription
End Sub

Public Property Get IsActionMenu() As Boolean
    IsActionMenu = mSelectionMode
End Property

Private Sub BindDesignerControls()
    Set mSearchText = FindDesignerControl("txt_search")
    Set mMenuEnrollment = FindDesignerControl("menuEnrollment")
    Set mMenuTransfer = FindDesignerControl("menuTransfer")
    Set mMenuExclusion = FindDesignerControl("menuExclusion")
    Set mMenuHistory = FindDesignerControl("menuHistory")
    Set mMenuClose = FindDesignerControl("menuClose")
    Set mPreviewConfirm = FindDesignerControl("btnPreviewConfirm")
    Set mPreviewCancel = FindDesignerControl("btnPreviewCancel")
    ApplyDesignerLocalization
End Sub

Private Sub ConfigureWizard()
    Dim actionType As String
    Dim multiPage As Object

    If mSelectionMode Then
        ConfigureActionMenu
        Exit Sub
    End If

    FindDesignerControl("fraActionMenu").Visible = False
    FindDesignerControl("fraWizard").Visible = True
    actionType = CurrentActionType()
    Set multiPage = FindDesignerControl("mpAction")
    multiPage.Style = fmTabStyleNone
    If actionType = "EXCLUSION" Then
        Me.Caption = t("personnel.wizard.title.exclusion", "Исключение из списков")
        multiPage.value = 1
    Else
        Me.Caption = t("personnel.wizard.title.transfer", "Кадровое перемещение")
        multiPage.value = 0
    End If
    ConfigureActionButtons
    LogDebug "configure-wizard", "event_type=" & actionType
End Sub

Private Sub ConfigureActionButtons()
    FindDesignerControl("btnExportRequest").Caption = t("personnel.wizard.find_load", "Найти и загрузить")
    FindDesignerControl("btnImportResponse").Caption = t("personnel.wizard.save", "Сохранить")
    FindDesignerControl("btnLicenseStatus").Caption = t("personnel.wizard.export", "Экспортировать Word")
    FindDesignerControl("btnClose").Caption = t("personnel.wizard.close", "Закрыть")
    FindDesignerControl("btnImportResponse").Enabled = True
    FindDesignerControl("btnExportRequest").Enabled = True
    FindDesignerControl("btnLicenseStatus").Enabled = (mSavedEventID <> "")
    FindDesignerControl("btnPreviewConfirm").Enabled = False
    FindDesignerControl("btnPreviewCancel").Enabled = False
End Sub

Private Sub ConfigureActionMenu()
    FindDesignerControl("fraWizard").Visible = False
    FindDesignerControl("fraActionMenu").Visible = True
    Me.Caption = t("ribbon.ui.personnelActionsGroup.label", "Кадровые действия")
    FindDesignerControl("lblDescription").Caption = t("ribbon.ui.personnelActionsGroup.label", "Кадровые действия")
    FindDesignerControl("menuEnrollment").Caption = t("ribbon.ui.openPersonnelEnrollmentAction.label", "Зачисление")
    FindDesignerControl("menuTransfer").Caption = t("ribbon.ui.openPersonnelTransferAction.label", "Перемещение")
    FindDesignerControl("menuExclusion").Caption = t("ribbon.ui.openPersonnelExclusionAction.label", "Исключение")
    FindDesignerControl("menuHistory").Caption = t("ribbon.ui.openPersonnelHistoryAction.label", "История сотрудника")
    FindDesignerControl("menuClose").Caption = t("personnel.wizard.close", "Закрыть")
    LogDebug "configure-menu", "ready=true"
End Sub

Private Sub ApplyDesignerLocalization()
    SetDesignerCaption "lbl_section_employee", "personnel.wizard.section.employee", "1. Найти сотрудника"
    SetDesignerCaption "lbl_search", "personnel.wizard.search", "Фамилия, личный или табельный номер"
    SetDesignerCaption "lbl_search_results", "personnel.wizard.search_results", "Результат поиска"
    SetDesignerCaption "lbl_section_order", "personnel.wizard.section.order", "2. Реквизиты действия"
    SetDesignerCaption "lbl_event_date", "personnel.wizard.event_date", "Дата события"
    SetDesignerCaption "lbl_effective_date", "personnel.wizard.effective_date", "Вступает в силу"
    SetDesignerCaption "lbl_order_reference", "personnel.wizard.order_reference", "Реквизиты приказа"
    SetDesignerCaption "lbl_basis_text", "personnel.wizard.basis", "Основание (войдёт в Word)"
    SetDesignerCaption "lbl_comment", "personnel.wizard.comment", "Служебный комментарий"
    SetDesignerCaption "lbl_section_transfer", "personnel.wizard.section.transfer", "3. Что меняется при перемещении"
    SetDesignerCaption "lbl_new_rank", "personnel.wizard.new_rank", "Новое звание"
    SetDesignerCaption "lbl_new_vus", "personnel.wizard.new_vus", "Новый ВУС"
    SetDesignerCaption "lbl_new_position", "personnel.wizard.new_position", "Новая должность"
    SetDesignerCaption "lbl_new_section", "personnel.wizard.new_section", "Подразделение"
    SetDesignerCaption "lbl_new_military_unit", "personnel.wizard.new_military_unit", "Воинская часть"
    SetDesignerCaption "lbl_section_dates", "personnel.wizard.section.dates", "4. Даты и место убытия"
    SetDesignerCaption "lbl_transfer_handover_date", "personnel.wizard.handover_date", "Сдал дела"
    SetDesignerCaption "lbl_acceptance_date", "personnel.wizard.acceptance_date", "Принял дела"
    SetDesignerCaption "lbl_duty_start_date", "personnel.wizard.duty_start_date", "Приступил"
    SetDesignerCaption "lbl_transfer_destination_unit", "personnel.wizard.destination_unit", "Куда убывает"
    SetDesignerCaption "lbl_transfer_destination_location", "personnel.wizard.destination_location", "Населённый пункт"
    SetDesignerCaption "lbl_section_exclusion", "personnel.wizard.section.exclusion", "3. Сведения об исключении"
    SetDesignerCaption "lbl_exclusion_handover_date", "personnel.wizard.handover_date", "Дата сдачи дел"
    SetDesignerCaption "lbl_exclusion_destination_unit", "personnel.wizard.destination_unit", "Куда убывает"
    SetDesignerCaption "lbl_exclusion_destination_location", "personnel.wizard.destination_location", "Населённый пункт"
    SetDesignerCaption "lbl_material_assistance_status", "personnel.wizard.material_assistance_status", "Материальная помощь за год"
    SetDesignerCaption "lbl_main_leave_status", "personnel.wizard.main_leave_status", "Основной отпуск за год"
    SetDesignerCaption "lbl_additional_leave_status", "personnel.wizard.additional_leave_status", "Дополнительный отпуск за год"
    SetDesignerCaption "pgTransfer", "ribbon.ui.openPersonnelTransferAction.label", "Перемещение"
    SetDesignerCaption "pgExclusion", "ribbon.ui.openPersonnelExclusionAction.label", "Исключение"
    SetDesignerCaption "pgPreview", "personnel.preview.page", "Проверка"
    SetDesignerCaption "lbl_preview_title", "personnel.preview.title", "5. Проверка перед сохранением"
    SetDesignerCaption "lbl_preview_before", "personnel.preview.before", "До"
    SetDesignerCaption "lbl_preview_after", "personnel.preview.after", "После"
    SetDesignerCaption "lbl_preview_payments", "personnel.preview.payments", "Выплаты"
    SetDesignerCaption "lbl_preview_warnings", "personnel.preview.warnings", "Предупреждения"
    SetDesignerCaption "btnPreviewConfirm", "personnel.preview.confirm", "Подтвердить"
    SetDesignerCaption "btnPreviewCancel", "personnel.preview.cancel", "Отмена"
End Sub

Private Sub SetDesignerCaption(ByVal controlName As String, ByVal localizationKey As String, ByVal fallbackText As String)
    FindDesignerControl(controlName).Caption = t(localizationKey, fallbackText)
End Sub

Private Function FindDesignerControl(ByVal controlName As String) As Object
    Set FindDesignerControl = FindDesignerControlInContainer(Me, controlName)
    If FindDesignerControl Is Nothing Then Err.Raise 5, "frmPersonnelActionWizardV2.FindDesignerControl", "Не найден design-time контрол: " & controlName
End Function

Private Function FindDesignerControlInContainer(ByVal containerHost As Object, ByVal controlName As String) As Object
    Dim foundControl As Object
    Dim controlItem As Object
    Dim pageItem As Object

    On Error Resume Next
    Set foundControl = containerHost.Controls.item(controlName)
    On Error GoTo 0
    If Not foundControl Is Nothing Then
        Set FindDesignerControlInContainer = foundControl
        Exit Function
    End If

    For Each controlItem In containerHost.Controls
        If typeName(controlItem) = "MultiPage" Then
            For Each pageItem In controlItem.Pages
                If StrComp(CStr(pageItem.Name), controlName, vbTextCompare) = 0 Then
                    Set FindDesignerControlInContainer = pageItem
                    Exit Function
                End If
                Set foundControl = FindDesignerControlInContainer(pageItem, controlName)
                If Not foundControl Is Nothing Then
                    Set FindDesignerControlInContainer = foundControl
                    Exit Function
                End If
            Next pageItem
        ElseIf typeName(controlItem) = "Frame" Then
            Set foundControl = FindDesignerControlInContainer(controlItem, controlName)
            If Not foundControl Is Nothing Then
                Set FindDesignerControlInContainer = foundControl
                Exit Function
            End If
        End If
    Next controlItem
End Function

Private Function CurrentActionType() As String
    CurrentActionType = UCase$(Trim$(CStr(mdlPersonnelEvents.GetPersonnelWizardValue("event_type"))))
End Function

Private Function DesignerNameForField(ByVal fieldKey As String) As String
    Select Case fieldKey
        Case "handover_date", "destination_unit", "destination_location"
            If CurrentActionType() = "EXCLUSION" Then
                DesignerNameForField = "txt_exclusion_" & fieldKey
            Else
                DesignerNameForField = "txt_transfer_" & fieldKey
            End If
        Case Else
            DesignerNameForField = "txt_" & fieldKey
    End Select
End Function

Private Sub mSearchText_Change()
    PreviewEmployeeSearch
End Sub

Private Sub PreviewEmployeeSearch()
    Dim item As Object
    Dim description As String
    Dim resultText As String
    Dim query As String

    query = Trim$(mSearchText.value)
    If Len(query) < 2 Then
        SetText "search_results", ""
        Exit Sub
    End If

    Set mSearchMatches = mdlPersonnelEvents.SearchPersonnelEmployees(query)
    For Each item In mSearchMatches
        description = CStr(item("fio")) & " - LN: " & CStr(item("personal_number")) & "; tab.: " & CStr(item("table_number"))
        If resultText <> "" Then resultText = resultText & " | "
        resultText = resultText & description
    Next item
    SetText "search_results", resultText
End Sub

Private Function FindAndLoadEmployee() As Boolean
    Dim item As Object
    Dim firstMatch As Object
    Dim description As String
    Dim resultText As String

    Set mSearchMatches = mdlPersonnelEvents.SearchPersonnelEmployees(TextOf("search"))
    For Each item In mSearchMatches
        description = CStr(item("fio")) & " — ЛН: " & CStr(item("personal_number")) & "; таб.: " & CStr(item("table_number"))
        If resultText <> "" Then resultText = resultText & " | "
        resultText = resultText & description
    Next item
    If mSearchMatches.count = 1 Then
        Set firstMatch = mSearchMatches(1)
        SetText "search_results", resultText
        SetText "employee_id", CStr(firstMatch("employee_id"))
        If mdlPersonnelEvents.LoadPersonnelWizardCurrentState() Then
            LoadValues
            mLoadedSignature = CurrentSignature
            SetText "status", t("personnel.wizard.employee_loaded", "Карточка сотрудника загружена.")
            FindAndLoadEmployee = True
        End If
    ElseIf mSearchMatches.count = 0 And Trim$(TextOf("search")) <> "" Then
        SetText "search_results", ""
        SetText "status", t("personnel.wizard.search_empty", "Сотрудник не найден.")
    ElseIf mSearchMatches.count > 1 Then
        SetText "search_results", resultText
        SetText "status", t("personnel.wizard.search_refine", "Несколько совпадений: уточните поиск личным или табельным номером.")
    End If
End Function

Private Sub LoadValues()
    Dim fieldKey As Variant
    For Each fieldKey In VisibleFieldKeys
        SetText CStr(fieldKey), PV(CStr(fieldKey))
    Next fieldKey
    mSavedEventID = PV("saved_event_id")
    FindDesignerControl("btnLicenseStatus").Enabled = (mSavedEventID <> "")
    If mSavedEventID <> "" Then
        mPreviewState = "SAVED"
        SetText "status", t("personnel.wizard.saved_prefix", "Сохранено:") & " " & mSavedEventID
    Else
        mPreviewState = "EDITING"
    End If
End Sub

Private Function VisibleFieldKeys() As Variant
    If CurrentActionType() = "EXCLUSION" Then
        VisibleFieldKeys = Array("employee_id", "event_date", "effective_date", "order_reference", "basis_text", "comment", "handover_date", "destination_unit", "destination_location", "material_assistance_status", "main_leave_status", "additional_leave_status", "status")
    Else
        VisibleFieldKeys = Array("employee_id", "event_date", "effective_date", "order_reference", "basis_text", "comment", "new_rank", "new_position", "new_section", "new_military_unit", "new_vus", "handover_date", "acceptance_date", "duty_start_date", "destination_unit", "destination_location", "status")
    End If
End Function

Private Function PV(ByVal fieldKey As String) As String
    Dim rawValue As Variant
    rawValue = mdlPersonnelEvents.GetPersonnelWizardValue(fieldKey)
    If IsDate(rawValue) Then PV = Format$(CDate(rawValue), "dd.mm.yyyy") Else PV = Trim$(CStr(rawValue))
End Function

Private Sub SetText(ByVal fieldKey As String, ByVal valueText As String)
    Dim targetControl As Object
    Set targetControl = FindDesignerControl(DesignerNameForField(fieldKey))
    targetControl.value = valueText
End Sub

Private Function TextOf(ByVal fieldKey As String) As String
    TextOf = Trim$(CStr(FindDesignerControl(DesignerNameForField(fieldKey)).value))
End Function

Private Sub WriteValues()
    Dim fieldKey As Variant
    For Each fieldKey In VisibleFieldKeys
        If CStr(fieldKey) <> "status" Then mdlPersonnelEvents.SetPersonnelWizardValue CStr(fieldKey), TextOf(CStr(fieldKey))
    Next fieldKey
End Sub

Private Function CurrentSignature() As String
    Dim fieldKey As Variant
    For Each fieldKey In VisibleFieldKeys
        If CStr(fieldKey) <> "status" Then CurrentSignature = CurrentSignature & "|" & CStr(fieldKey) & "=" & TextOf(CStr(fieldKey))
    Next fieldKey
End Function

Public Function SaveAction() As String
    If mPreviewState <> "CONFIRMED" Then
        If mPreviewState = "PREVIEW_READY" Then
            SetText "status", t("personnel.preview.confirm_required", "Подтвердите просмотр перед сохранением действия.")
        Else
            Call PrepareConfirmationPreview
            SetText "status", t("personnel.preview.confirm_required", "Подтвердите просмотр перед сохранением действия.")
        End If
        Exit Function
    End If
    On Error GoTo Failed
    LogDebug "save-start", "event_type=" & CurrentActionType()
    WriteValues
    mSavedEventID = mdlPersonnelEvents.SavePersonnelWizardAction()
    mdlPersonnelEvents.SetPersonnelWizardValue "saved_event_id", mSavedEventID
    FindDesignerControl("btnLicenseStatus").Enabled = True
    SetText "status", t("personnel.wizard.saved_prefix", "Сохранено:") & " " & mSavedEventID
    mLoadedSignature = CurrentSignature
    mPreviewState = "SAVED"
    FindDesignerControl("btnImportResponse").Enabled = False
    FindDesignerControl("btnExportRequest").Enabled = False
    FindDesignerControl("btnLicenseStatus").Enabled = True
    FindDesignerControl("btnPreviewConfirm").Enabled = False
    FindDesignerControl("btnPreviewCancel").Enabled = False
    SaveAction = mSavedEventID
    LogDebug "save-complete", "event_id=" & mSavedEventID
    Exit Function
Failed:
    mPreviewState = "PREVIEW_READY"
    FindDesignerControl("btnImportResponse").Enabled = False
    FindDesignerControl("btnExportRequest").Enabled = False
    FindDesignerControl("btnPreviewConfirm").Enabled = CBool(mPreviewResult("can_confirm"))
    FindDesignerControl("btnPreviewCancel").Enabled = True
    LogDebug "save-error", "number=" & CStr(Err.number) & "; description=" & Err.description
    SetText "status", Err.description
    Application.StatusBar = Err.description
End Function

Public Function ExportAction() As String
    On Error GoTo Failed
    If mSavedEventID = "" Or mPreviewState <> "SAVED" Then
        SetText "status", t("personnel.wizard.export_after_save", "Сначала сохраните кадровое действие.")
        Exit Function
    End If
    LogDebug "export-start", "event_id=" & mSavedEventID
    ExportAction = mdlPersonnelEventOrderExport.ExportPersonnelEventOrder(mSavedEventID)
    SetText "status", ExportAction
    LogDebug "export-complete", "event_id=" & mSavedEventID
    Exit Function
Failed:
    LogDebug "export-error", "event_id=" & mSavedEventID & "; number=" & CStr(Err.number) & "; description=" & Err.description
    SetText "status", Err.description
    Application.StatusBar = Err.description
End Function

Private Sub btnExportRequest_Click()
    If mPreviewState <> "EDITING" And mPreviewState <> "SAVED" Then CancelPreviewInternal False
    If Trim$(TextOf("search")) <> "" Then
        Call FindAndLoadEmployee
    Else
        WriteValues
        If mdlPersonnelEvents.LoadPersonnelWizardCurrentState() Then
            LoadValues
            mLoadedSignature = CurrentSignature
            SetText "status", t("personnel.wizard.employee_loaded", "Карточка сотрудника загружена.")
        Else
            SetText "status", CStr(Application.StatusBar)
        End If
    End If
End Sub

Private Sub btnImportResponse_Click()
    Call PrepareConfirmationPreview
End Sub

Private Sub btnLicenseStatus_Click()
    Call ExportAction
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mSelectionMode Then Exit Sub
    If CurrentSignature <> mLoadedSignature Then
        If MsgBox(t("personnel.wizard.unsaved_prompt", "Есть несохранённые изменения. Закрыть без сохранения?"), vbExclamation + vbYesNo) = vbNo Then Cancel = True
    End If
End Sub

Private Sub mMenuEnrollment_Click()
    Unload Me
    mdlPersonnelEvents.OpenPersonnelEnrollmentAction
End Sub

Private Sub mMenuTransfer_Click()
    Unload Me
    mdlPersonnelEvents.OpenPersonnelTransferAction
End Sub

Private Sub mMenuExclusion_Click()
    Unload Me
    mdlPersonnelEvents.OpenPersonnelExclusionAction
End Sub

Private Sub mMenuHistory_Click()
    Unload Me
    mdlPersonnelHistory.OpenPersonnelHistory
End Sub

Private Sub mMenuClose_Click()
    Unload Me
End Sub

Private Sub mPreviewConfirm_Click()
    Call ConfirmPreviewInternal
End Sub

Private Sub mPreviewCancel_Click()
    CancelPreviewInternal True
End Sub

Public Function PrepareConfirmationPreview() As Boolean
    Dim draft As Object
    Dim preview As Object

    On Error GoTo Failed
    If mSelectionMode Then Exit Function
    Set draft = CollectCurrentDraft()
    Set preview = mdlPersonnelActionPreview.BuildPersonnelActionPreview(draft)
    Set mPreviewDraft = draft
    Set mPreviewResult = preview
    mPreviewSignature = CurrentSignature
    RenderPreview preview
    FindDesignerControl("mpAction").value = 2
    mPreviewState = "PREVIEW_READY"
    FindDesignerControl("btnImportResponse").Enabled = False
    FindDesignerControl("btnExportRequest").Enabled = False
    FindDesignerControl("btnLicenseStatus").Enabled = False
    FindDesignerControl("btnPreviewConfirm").Enabled = CBool(preview("can_confirm"))
    FindDesignerControl("btnPreviewCancel").Enabled = True
    If CBool(preview("can_confirm")) Then
        SetText "status", t("personnel.preview.ready", "Проверка готова. Проверьте данные и подтвердите.")
    Else
        SetText "status", t("personnel.preview.invalid", "Предпросмотр содержит ошибки. Сохранение недоступно.")
    End If
    LogDebug "confirmation-opened", "page=preview"
    PrepareConfirmationPreview = CBool(preview("can_confirm"))
    Exit Function
Failed:
    SetText "status", Err.description
    Application.StatusBar = Err.description
    LogDebug "confirmation-open-error", "number=" & CStr(Err.number)
End Function

Public Sub OpenConfirmationPreview()
    Call PrepareConfirmationPreview
End Sub

Public Function ConfirmConfirmationPreview() As String
    ConfirmConfirmationPreview = ConfirmPreviewInternal()
End Function

Public Sub CancelConfirmationPreview()
    CancelPreviewInternal True
End Sub

Private Function ConfirmPreviewInternal() As String
    If mPreviewState = "SAVED" Then
        LogDebug "confirmation-confirmed-ignored", "state=saved"
        ConfirmPreviewInternal = mSavedEventID
        Exit Function
    End If
    If mPreviewState <> "PREVIEW_READY" Then
        SetText "status", t("personnel.preview.confirm_required", "Подтвердите просмотр перед сохранением действия.")
        Exit Function
    End If
    If mPreviewResult Is Nothing Then
        SetText "status", t("personnel.preview.confirm_required", "Подтвердите просмотр перед сохранением действия.")
        Exit Function
    End If
    If CurrentSignature <> mPreviewSignature Then
        CancelPreviewInternal False
        SetText "status", t("personnel.preview.changed", "Черновик изменился. Предпросмотр сброшен; проверьте его снова.")
        Exit Function
    End If
    If Not CBool(mPreviewResult("can_confirm")) Then
        SetText "status", t("personnel.preview.invalid", "Предпросмотр содержит ошибки. Сохранение недоступно.")
        Exit Function
    End If
    mPreviewState = "CONFIRMED"
    LogDebug "confirmation-confirmed", "page=preview"
    ConfirmPreviewInternal = SaveAction()
End Function

Private Sub CancelPreviewInternal(ByVal showStatus As Boolean)
    Set mPreviewDraft = Nothing
    Set mPreviewResult = Nothing
    mPreviewSignature = ""
    If Not mSelectionMode Then
        FindDesignerControl("mpAction").value = IIf(CurrentActionType() = "EXCLUSION", 1, 0)
        mPreviewState = "EDITING"
        FindDesignerControl("btnImportResponse").Enabled = True
        FindDesignerControl("btnExportRequest").Enabled = True
        FindDesignerControl("btnLicenseStatus").Enabled = (mSavedEventID <> "")
        FindDesignerControl("btnPreviewConfirm").Enabled = False
        FindDesignerControl("btnPreviewCancel").Enabled = False
    End If
    ClearPreviewControls
    If showStatus Then SetText "status", t("personnel.preview.cancelled", "Предпросмотр отменён. Черновик не сохранён.")
    LogDebug "confirmation-cancelled", "page=preview"
End Sub

Private Function CollectCurrentDraft() As Object
    Dim draft As Object
    Dim fieldKey As Variant

    Set draft = CreateObject("Scripting.Dictionary")
    draft.CompareMode = vbTextCompare
    draft.Add "event_type", CurrentActionType()
    For Each fieldKey In VisibleFieldKeys
        If CStr(fieldKey) <> "status" Then draft(CStr(fieldKey)) = TextOf(CStr(fieldKey))
    Next fieldKey
    Set CollectCurrentDraft = draft
End Function

Private Sub RenderPreview(ByVal preview As Object)
    FindDesignerControl("txt_preview_before").value = RenderChangedLines(preview("changed_fields"), "before")
    FindDesignerControl("txt_preview_after").value = RenderChangedLines(preview("changed_fields"), "after")
    FindDesignerControl("txt_preview_payments").value = RenderPaymentLines(preview("payment_changes"))
    FindDesignerControl("txt_preview_warnings").value = RenderWarningLines(preview("warnings"))
End Sub

Private Sub ClearPreviewControls()
    FindDesignerControl("txt_preview_before").value = ""
    FindDesignerControl("txt_preview_after").value = ""
    FindDesignerControl("txt_preview_payments").value = ""
    FindDesignerControl("txt_preview_warnings").value = ""
End Sub

Private Function RenderChangedLines(ByVal changedFields As Collection, ByVal valueKey As String) As String
    Dim item As Object
    Dim lineText As String
    Dim valueText As String

    For Each item In changedFields
        If UCase$(PreviewValue(item("change_kind"))) <> "UNCHANGED" Then
            valueText = PreviewValue(item(valueKey))
            lineText = PreviewFieldLabel(CStr(item("key"))) & ": " & valueText
            If RenderChangedLines <> "" Then RenderChangedLines = RenderChangedLines & vbCrLf
            RenderChangedLines = RenderChangedLines & lineText
        End If
    Next item
    If RenderChangedLines = "" Then RenderChangedLines = t("personnel.preview.no_changes", "Изменений нет.")
End Function

Private Function RenderPaymentLines(ByVal paymentChanges As Collection) As String
    Dim item As Object
    Dim lineText As String
    Dim amountText As String

    For Each item In paymentChanges
        lineText = UCase$(PreviewValue(item("change_kind"))) & ": " & PreviewValue(item("payment_code"))
        amountText = PreviewValue(item("amount_value"), "")
        If amountText <> "" Then lineText = lineText & " (" & amountText & ")"
        If RenderPaymentLines <> "" Then RenderPaymentLines = RenderPaymentLines & vbCrLf
        RenderPaymentLines = RenderPaymentLines & lineText
    Next item
    If RenderPaymentLines = "" Then RenderPaymentLines = t("personnel.preview.no_payments", "Изменений выплат нет.")
End Function

Private Function RenderWarningLines(ByVal warnings As Collection) As String
    Dim item As Object
    Dim lineText As String

    For Each item In warnings
        lineText = UCase$(PreviewValue(item("severity"))) & " " & PreviewValue(item("code")) & ": " & PreviewValue(item("detail"))
        If RenderWarningLines <> "" Then RenderWarningLines = RenderWarningLines & vbCrLf
        RenderWarningLines = RenderWarningLines & lineText
    Next item
    If RenderWarningLines = "" Then RenderWarningLines = t("personnel.preview.no_warnings", "Предупреждений нет.")
End Function

Private Function PreviewFieldLabel(ByVal fieldKey As String) As String
    Select Case fieldKey
        Case "new_rank": PreviewFieldLabel = CStr(FindDesignerControl("lbl_new_rank").Caption)
        Case "new_position": PreviewFieldLabel = CStr(FindDesignerControl("lbl_new_position").Caption)
        Case "new_section": PreviewFieldLabel = CStr(FindDesignerControl("lbl_new_section").Caption)
        Case "new_military_unit": PreviewFieldLabel = CStr(FindDesignerControl("lbl_new_military_unit").Caption)
        Case "new_vus": PreviewFieldLabel = CStr(FindDesignerControl("lbl_new_vus").Caption)
        Case "event_date": PreviewFieldLabel = CStr(FindDesignerControl("lbl_event_date").Caption)
        Case "effective_date": PreviewFieldLabel = CStr(FindDesignerControl("lbl_effective_date").Caption)
        Case "order_reference": PreviewFieldLabel = CStr(FindDesignerControl("lbl_order_reference").Caption)
        Case "basis_text": PreviewFieldLabel = CStr(FindDesignerControl("lbl_basis_text").Caption)
        Case "comment": PreviewFieldLabel = CStr(FindDesignerControl("lbl_comment").Caption)
        Case "handover_date": PreviewFieldLabel = CStr(FindDesignerControl(IIf(CurrentActionType() = "EXCLUSION", "lbl_exclusion_handover_date", "lbl_transfer_handover_date")).Caption)
        Case "acceptance_date": PreviewFieldLabel = CStr(FindDesignerControl("lbl_acceptance_date").Caption)
        Case "duty_start_date": PreviewFieldLabel = CStr(FindDesignerControl("lbl_duty_start_date").Caption)
        Case "destination_unit": PreviewFieldLabel = CStr(FindDesignerControl(IIf(CurrentActionType() = "EXCLUSION", "lbl_exclusion_destination_unit", "lbl_transfer_destination_unit")).Caption)
        Case "destination_location": PreviewFieldLabel = CStr(FindDesignerControl(IIf(CurrentActionType() = "EXCLUSION", "lbl_exclusion_destination_location", "lbl_transfer_destination_location")).Caption)
        Case "material_assistance_status": PreviewFieldLabel = CStr(FindDesignerControl("lbl_material_assistance_status").Caption)
        Case "main_leave_status": PreviewFieldLabel = CStr(FindDesignerControl("lbl_main_leave_status").Caption)
        Case "additional_leave_status": PreviewFieldLabel = CStr(FindDesignerControl("lbl_additional_leave_status").Caption)
        Case Else: PreviewFieldLabel = fieldKey
    End Select
End Function

Private Function PreviewValue(ByVal rawValue As Variant, Optional ByVal emptyText As String = "-") As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    If IsDate(rawValue) Then
        PreviewValue = Format$(CDate(rawValue), "dd.mm.yyyy")
    Else
        PreviewValue = Trim$(CStr(rawValue))
    End If
    If PreviewValue = "" Then PreviewValue = emptyText
End Function

Private Sub LogDebug(ByVal actionName As String, ByVal detailText As String)
    If DEBUG_LOGGING Then Debug.Print "[PERSONNEL-ACTION-V2] action=" & actionName & "; " & detailText
End Sub
