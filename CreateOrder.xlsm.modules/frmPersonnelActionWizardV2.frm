VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPersonnelActionWizardV2 
   Caption         =   "Кадровое действие V2"
   ClientHeight    =   10620
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

Private Sub UserForm_Initialize()
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    mSelectionMode = mdlPersonnelEvents.ConsumePersonnelActionMenuRequest()
    LogDebug "initialize", "selection_mode=" & CStr(mSelectionMode)
    BindDesignerControls
    ConfigureWizard
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
    If mSavedEventID <> "" Then SetText "status", t("personnel.wizard.saved_prefix", "Сохранено:") & " " & mSavedEventID
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
    On Error GoTo Failed
    LogDebug "save-start", "event_type=" & CurrentActionType()
    WriteValues
    mSavedEventID = mdlPersonnelEvents.SavePersonnelWizardAction()
    mdlPersonnelEvents.SetPersonnelWizardValue "saved_event_id", mSavedEventID
    FindDesignerControl("btnLicenseStatus").Enabled = True
    SetText "status", t("personnel.wizard.saved_prefix", "Сохранено:") & " " & mSavedEventID
    mLoadedSignature = CurrentSignature
    SaveAction = mSavedEventID
    LogDebug "save-complete", "event_id=" & mSavedEventID
    Exit Function
Failed:
    LogDebug "save-error", "number=" & CStr(Err.number) & "; description=" & Err.description
    SetText "status", Err.description
    Application.StatusBar = Err.description
End Function

Public Function ExportAction() As String
    On Error GoTo Failed
    If mSavedEventID = "" Then
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
    Call SaveAction
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

Private Sub LogDebug(ByVal actionName As String, ByVal detailText As String)
    If DEBUG_LOGGING Then Debug.Print "[PERSONNEL-ACTION-V2] action=" & actionName & "; " & detailText
End Sub
