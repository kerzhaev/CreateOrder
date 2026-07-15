VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPersonnelActionWizard
   Caption         =   "Кадровое действие"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmPersonnelActionWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPersonnelActionWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSavedEventID As String
Private mSelectionMode As Boolean
Private mLoadedSignature As String
Private mSearchMatches As Collection
Private WithEvents mSearchText As MSForms.TextBox
Private WithEvents mMenuEnrollment As MSForms.CommandButton
Private WithEvents mMenuTransfer As MSForms.CommandButton
Private WithEvents mMenuExclusion As MSForms.CommandButton
Private WithEvents mMenuHistory As MSForms.CommandButton
Private WithEvents mMenuClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    mSelectionMode = mdlPersonnelEvents.ConsumePersonnelActionMenuRequest()
    ConfigureWizard
    If Not mSelectionMode Then
        LoadValues
        mLoadedSignature = CurrentSignature
    End If
End Sub

Public Property Get IsActionMenu() As Boolean
    IsActionMenu = mSelectionMode
End Property

Private Sub ConfigureWizard()
    Dim actionType As String

    If mSelectionMode Then
        ConfigureActionMenu
        Exit Sub
    End If

    lblDescription.Visible = False
    actionType = UCase$(CStr(mdlPersonnelEvents.GetPersonnelWizardValue("event_type")))
    If actionType = "EXCLUSION" Then
        Me.Caption = t("personnel.wizard.title.exclusion", "Исключение из списков")
    Else
        Me.Caption = t("personnel.wizard.title.transfer", "Кадровое перемещение")
    End If
    Me.ScrollBars = fmScrollBarsNone
    Me.Width = 790
    Me.Height = IIf(actionType = "EXCLUSION", 470, 560)

    ConfigureActionButtons
    AddSearchBlock
    AddOrderBlock
    If actionType = "EXCLUSION" Then
        AddExclusionBlock
    Else
        AddTransferBlock
    End If
End Sub

Private Sub ConfigureActionButtons()
    btnExportRequest.Caption = t("personnel.wizard.find_load", "Найти и загрузить")
    btnImportResponse.Caption = t("personnel.wizard.save", "Сохранить")
    btnLicenseStatus.Caption = t("personnel.wizard.export", "Экспортировать Word")
    btnClose.Caption = t("personnel.wizard.close", "Закрыть")
    btnExportRequest.Left = 18: btnImportResponse.Left = 175: btnLicenseStatus.Left = 310: btnClose.Left = 575
    btnExportRequest.Top = Me.InsideHeight - 38
    btnImportResponse.Top = Me.InsideHeight - 38
    btnLicenseStatus.Top = Me.InsideHeight - 38
    btnClose.Top = Me.InsideHeight - 38
    btnExportRequest.Width = 145: btnImportResponse.Width = 120: btnLicenseStatus.Width = 180: btnClose.Width = 110
End Sub

Private Sub AddSearchBlock()
    Dim searchLabel As Object
    Dim resultLabel As Object
    Dim searchText As MSForms.TextBox
    Dim resultText As Object
    Dim employeeIdText As Object

    AddSectionTitle "section_employee", t("personnel.wizard.section.employee", "1. Найти сотрудника"), 16
    Set searchLabel = AddLabel("lbl_search", t("personnel.wizard.search", "Фамилия, личный или табельный номер"), 18, 38, 220)
    Set searchText = Me.Controls.Add("Forms.TextBox.1", "txt_search", True)
    searchText.Left = 242: searchText.Top = 36: searchText.Width = 400: searchText.Height = 20
    Set mSearchText = searchText
    Set resultLabel = AddLabel("lbl_search_results", t("personnel.wizard.search_results", "Результат поиска"), 18, 64, 110)
    Set resultText = Me.Controls.Add("Forms.TextBox.1", "txt_search_results", True)
    resultText.Left = 130: resultText.Top = 62: resultText.Width = 620: resultText.Height = 20
    resultText.Locked = True
    resultText.BackColor = RGB(242, 242, 242)

    'Идентификатор хранится технически: оператор его не вводит и не видит.
    Set employeeIdText = Me.Controls.Add("Forms.TextBox.1", "txt_employee_id", True)
    employeeIdText.Visible = False
End Sub

Private Sub AddOrderBlock()
    AddSectionTitle "section_order", t("personnel.wizard.section.order", "2. Реквизиты действия"), 94
    AddField "event_date", t("personnel.wizard.event_date", "Дата события"), 18, 116, 95, 100, False
    AddField "effective_date", t("personnel.wizard.effective_date", "Вступает в силу"), 235, 116, 105, 100, False
    AddField "order_reference", t("personnel.wizard.order_reference", "Реквизиты приказа"), 462, 116, 118, 170, False
    AddField "basis_text", t("personnel.wizard.basis", "Основание (войдёт в Word)"), 18, 144, 170, 560, True
    AddField "comment", t("personnel.wizard.comment", "Служебный комментарий"), 18, 194, 160, 570, True
End Sub

Private Sub AddTransferBlock()
    AddSectionTitle "section_transfer", t("personnel.wizard.section.transfer", "3. Что меняется при перемещении"), 246
    AddField "new_rank", t("personnel.wizard.new_rank", "Новое звание"), 18, 268, 100, 180, False
    AddField "new_vus", t("personnel.wizard.new_vus", "Новый ВУС"), 370, 268, 80, 180, False
    AddField "new_position", t("personnel.wizard.new_position", "Новая должность"), 18, 296, 110, 620, False
    AddField "new_section", t("personnel.wizard.new_section", "Подразделение"), 18, 324, 100, 220, False
    AddField "new_military_unit", t("personnel.wizard.new_military_unit", "Воинская часть"), 370, 324, 100, 260, False
    AddSectionTitle "section_dates", t("personnel.wizard.section.dates", "4. Даты и место убытия"), 356
    AddField "handover_date", t("personnel.wizard.handover_date", "Сдал дела"), 18, 378, 70, 100, False
    AddField "acceptance_date", t("personnel.wizard.acceptance_date", "Принял дела"), 220, 378, 78, 100, False
    AddField "duty_start_date", t("personnel.wizard.duty_start_date", "Приступил"), 420, 378, 72, 100, False
    AddField "destination_unit", t("personnel.wizard.destination_unit", "Куда убывает"), 18, 406, 100, 260, False
    AddField "destination_location", t("personnel.wizard.destination_location", "Населённый пункт"), 400, 406, 105, 230, False
    AddStatusField 438
End Sub

Private Sub AddExclusionBlock()
    AddSectionTitle "section_exclusion", t("personnel.wizard.section.exclusion", "3. Сведения об исключении"), 246
    AddField "handover_date", t("personnel.wizard.handover_date", "Дата сдачи дел"), 18, 268, 95, 110, False
    AddField "destination_unit", t("personnel.wizard.destination_unit", "Куда убывает"), 250, 268, 100, 280, False
    AddField "destination_location", t("personnel.wizard.destination_location", "Населённый пункт"), 18, 296, 105, 250, False
    AddStatusField 330
End Sub

Private Sub AddStatusField(ByVal topValue As Single)
    AddField "status", "", 18, topValue, 1, 700, False
    Me.Controls("txt_status").Locked = True
    Me.Controls("txt_status").BackColor = RGB(242, 242, 242)
End Sub

Private Sub AddSectionTitle(ByVal controlName As String, ByVal captionText As String, ByVal topValue As Single)
    Dim labelControl As Object
    Set labelControl = AddLabel(controlName, captionText, 18, topValue, 700)
    labelControl.Font.Bold = True
    labelControl.Font.Size = 10
End Sub

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single) As Object
    Dim result As Object
    Set result = Me.Controls.Add("Forms.Label.1", controlName, True)
    result.Caption = captionText
    result.Left = leftValue
    result.Top = topValue
    result.Width = widthValue
    result.Height = 18
    Set AddLabel = result
End Function

Private Sub AddField(ByVal fieldKey As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal labelWidth As Single, ByVal inputWidth As Single, ByVal multiline As Boolean)
    Dim labelControl As Object
    Dim textControl As Object

    If captionText <> "" Then Set labelControl = AddLabel("lbl_" & fieldKey, captionText, leftValue, topValue, labelWidth)
    Set textControl = Me.Controls.Add("Forms.TextBox.1", "txt_" & fieldKey, True)
    textControl.Left = leftValue + labelWidth + 8
    textControl.Top = topValue - 2
    textControl.Width = inputWidth
    If multiline Then
        textControl.Height = 38
        textControl.MultiLine = True
        textControl.EnterKeyBehavior = True
        textControl.ScrollBars = fmScrollBarsVertical
    Else
        textControl.Height = 20
    End If
End Sub

Private Sub ConfigureActionMenu()
    HideLegacyActionButtons
    Me.Caption = t("ribbon.ui.personnelActionsGroup.label", "Кадровые действия")
    Me.ScrollBars = fmScrollBarsNone
    Me.Width = 520: Me.Height = 300
    lblDescription.Visible = True
    lblDescription.Caption = t("ribbon.ui.personnelActionsGroup.label", "Кадровые действия")
    lblDescription.Left = 24: lblDescription.Top = 24: lblDescription.Width = 420: lblDescription.Height = 24
    lblDescription.Font.Bold = True
    Set mMenuEnrollment = AddMenuButton("menuEnrollment", t("ribbon.ui.openPersonnelEnrollmentAction.label", "Зачисление"), 24, 66)
    Set mMenuTransfer = AddMenuButton("menuTransfer", t("ribbon.ui.openPersonnelTransferAction.label", "Перемещение"), 24, 108)
    Set mMenuExclusion = AddMenuButton("menuExclusion", t("ribbon.ui.openPersonnelExclusionAction.label", "Исключение"), 24, 150)
    Set mMenuHistory = AddMenuButton("menuHistory", t("ribbon.ui.openPersonnelHistoryAction.label", "История сотрудника"), 24, 192)
    Set mMenuClose = AddMenuButton("menuClose", t("personnel.wizard.close", "Закрыть"), 316, 232)
    mMenuClose.Width = 128
End Sub

Private Sub HideLegacyActionButtons()
    btnExportRequest.Visible = False: btnImportResponse.Visible = False
    btnLicenseStatus.Visible = False: btnClose.Visible = False
End Sub

Private Function AddMenuButton(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single) As MSForms.CommandButton
    Dim result As MSForms.CommandButton
    Set result = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    result.Caption = captionText
    result.Left = leftValue: result.Top = topValue
    result.Width = 420: result.Height = 28
    Set AddMenuButton = result
End Function

Private Sub mSearchText_Change()
    PreviewEmployeeSearch
End Sub

Private Sub PreviewEmployeeSearch()
    Dim item As Object
    Dim description As String
    Dim resultText As String
    Dim query As String

    query = Trim$(mSearchText.Value)
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
    If mSearchMatches.Count = 1 Then
        Set firstMatch = mSearchMatches(1)
        SetText "search_results", resultText
        SetText "employee_id", CStr(firstMatch("employee_id"))
        If mdlPersonnelEvents.LoadPersonnelWizardCurrentState() Then
            LoadValues
            mLoadedSignature = CurrentSignature
            SetText "status", t("personnel.wizard.employee_loaded", "Карточка сотрудника загружена.")
            FindAndLoadEmployee = True
        End If
    ElseIf mSearchMatches.Count = 0 And Trim$(TextOf("search")) <> "" Then
        SetText "search_results", ""
        SetText "status", t("personnel.wizard.search_empty", "Сотрудник не найден.")
    ElseIf mSearchMatches.Count > 1 Then
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
    btnLicenseStatus.Enabled = (mSavedEventID <> "")
    If mSavedEventID <> "" Then SetText "status", t("personnel.wizard.saved_prefix", "Сохранено:") & " " & mSavedEventID
End Sub

Private Function VisibleFieldKeys() As Variant
    If UCase$(CStr(mdlPersonnelEvents.GetPersonnelWizardValue("event_type"))) = "EXCLUSION" Then
        VisibleFieldKeys = Array("employee_id", "event_date", "effective_date", "order_reference", "basis_text", "comment", "handover_date", "destination_unit", "destination_location", "status")
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
    If ControlExists("txt_" & fieldKey) Then Me.Controls("txt_" & fieldKey).Value = valueText
End Sub

Private Function TextOf(ByVal fieldKey As String) As String
    If ControlExists("txt_" & fieldKey) Then TextOf = Trim$(CStr(Me.Controls("txt_" & fieldKey).Value))
End Function

Private Function ControlExists(ByVal controlName As String) As Boolean
    Dim ignored As Object
    On Error Resume Next
    Set ignored = Me.Controls(controlName)
    ControlExists = Not ignored Is Nothing
    On Error GoTo 0
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
    WriteValues
    mSavedEventID = mdlPersonnelEvents.SavePersonnelWizardAction()
    mdlPersonnelEvents.SetPersonnelWizardValue "saved_event_id", mSavedEventID
    btnLicenseStatus.Enabled = True
    SetText "status", t("personnel.wizard.saved_prefix", "Сохранено:") & " " & mSavedEventID
    mLoadedSignature = CurrentSignature
    SaveAction = mSavedEventID
    Exit Function
Failed:
    SetText "status", Err.Description
    Application.StatusBar = Err.Description
End Function

Public Function ExportAction() As String
    On Error GoTo Failed
    If mSavedEventID = "" Then
        SetText "status", t("personnel.wizard.export_after_save", "Сначала сохраните кадровое действие.")
        Exit Function
    End If
    ExportAction = mdlPersonnelEventOrderExport.ExportPersonnelEventOrder(mSavedEventID)
    SetText "status", ExportAction
    Exit Function
Failed:
    SetText "status", Err.Description
    Application.StatusBar = Err.Description
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
