VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPersonnelActionWizard
   Caption         =   "Файлы лицензии"
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
Private WithEvents mMenuEnrollment As MSForms.CommandButton
Private WithEvents mMenuTransfer As MSForms.CommandButton
Private WithEvents mMenuExclusion As MSForms.CommandButton
Private WithEvents mMenuHistory As MSForms.CommandButton
Private WithEvents mMenuClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    mSelectionMode = mdlPersonnelEvents.ConsumePersonnelActionMenuRequest()
    ConfigureWizard
    If Not mSelectionMode Then LoadValues
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
        Me.Caption = t("personnel.wizard.title.exclusion", "Exclusion from personnel list")
    Else
        Me.Caption = t("personnel.wizard.title.transfer", "Personnel transfer")
    End If
    Me.ScrollBars = fmScrollBarsVertical
    Me.ScrollHeight = 680
    Me.Width = 760
    Me.Height = 520

    btnExportRequest.Caption = t("personnel.wizard.load", "Load current state")
    btnImportResponse.Caption = t("personnel.wizard.save", "Save")
    btnLicenseStatus.Caption = t("personnel.wizard.export", "Export order")
    btnClose.Caption = t("personnel.wizard.close", "Close")
    btnExportRequest.Left = 18: btnImportResponse.Left = 170: btnLicenseStatus.Left = 322: btnClose.Left = 574
    btnExportRequest.Top = 625: btnImportResponse.Top = 625: btnLicenseStatus.Top = 625: btnClose.Top = 625

    AddField "employee_id", t("personnel.wizard.employee_id", "EmployeeID"), 18, False
    AddField "event_date", t("personnel.wizard.event_date", "Event date"), 46, False
    AddField "effective_date", t("personnel.wizard.effective_date", "Effective date"), 74, False
    AddField "order_reference", t("personnel.wizard.order_reference", "Order reference"), 102, False
    AddField "basis_text", t("personnel.wizard.basis", "Basis"), 130, True
    AddField "comment", t("personnel.wizard.comment", "Comment"), 184, True
    AddField "new_rank", t("personnel.wizard.new_rank", "New rank"), 238, False
    AddField "new_position", t("personnel.wizard.new_position", "New position"), 266, False
    AddField "new_section", t("personnel.wizard.new_section", "New section"), 294, False
    AddField "new_military_unit", t("personnel.wizard.new_military_unit", "New military unit"), 322, False
    AddField "new_vus", t("personnel.wizard.new_vus", "New VUS"), 350, False
    AddField "handover_date", t("personnel.wizard.handover_date", "Handover date"), 378, False
    AddField "acceptance_date", t("personnel.wizard.acceptance_date", "Acceptance date"), 406, False
    AddField "duty_start_date", t("personnel.wizard.duty_start_date", "Duty start date"), 434, False
    AddField "destination_unit", t("personnel.wizard.destination_unit", "Destination unit"), 462, False
    AddField "destination_location", t("personnel.wizard.destination_location", "Destination location"), 490, False
    AddField "status", "", 534, True
    Me.Controls("txt_status").Locked = True
End Sub

Private Sub ConfigureActionMenu()
    HideLegacyActionButtons
    Me.Caption = t("ribbon.ui.personnelActionsGroup.label", "Personnel actions")
    Me.ScrollBars = fmScrollBarsNone
    Me.ScrollHeight = 0
    Me.Width = 520
    Me.Height = 300
    lblDescription.Visible = True
    lblDescription.Caption = t("ribbon.ui.personnelActionsGroup.label", "Personnel actions")
    lblDescription.Left = 24
    lblDescription.Top = 24
    lblDescription.Width = 420
    lblDescription.Height = 24
    lblDescription.Font.Bold = True

    Set mMenuEnrollment = AddMenuButton("menuEnrollment", t("ribbon.ui.openPersonnelEnrollmentAction.label", "Enrollment"), 24, 66)
    Set mMenuTransfer = AddMenuButton("menuTransfer", t("ribbon.ui.openPersonnelTransferAction.label", "Transfer"), 24, 108)
    Set mMenuExclusion = AddMenuButton("menuExclusion", t("ribbon.ui.openPersonnelExclusionAction.label", "Exclusion"), 24, 150)
    Set mMenuHistory = AddMenuButton("menuHistory", t("ribbon.ui.openPersonnelHistoryAction.label", "Employee history"), 24, 192)
    Set mMenuClose = AddMenuButton("menuClose", t("personnel.wizard.close", "Close"), 316, 232)
    mMenuClose.Width = 128
End Sub

Private Sub HideLegacyActionButtons()
    btnExportRequest.Visible = False
    btnImportResponse.Visible = False
    btnLicenseStatus.Visible = False
    btnClose.Visible = False
End Sub

Private Function AddMenuButton(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single) As MSForms.CommandButton
    Dim buttonControl As MSForms.CommandButton

    Set buttonControl = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    buttonControl.Caption = captionText
    buttonControl.Left = leftValue
    buttonControl.Top = topValue
    buttonControl.Width = 420
    buttonControl.Height = 28
    Set AddMenuButton = buttonControl
End Function

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
Private Sub AddField(ByVal fieldKey As String, ByVal captionText As String, ByVal topValue As Single, ByVal multiline As Boolean)
    Dim labelControl As Object
    Dim textControl As Object

    Set labelControl = Me.Controls.Add("Forms.Label.1", "lbl_" & fieldKey, True)
    labelControl.Caption = captionText
    labelControl.Left = 18
    labelControl.Top = topValue
    labelControl.Width = 215
    labelControl.Height = 18

    Set textControl = Me.Controls.Add("Forms.TextBox.1", "txt_" & fieldKey, True)
    textControl.Left = 240
    textControl.Top = topValue - 2
    textControl.Width = 480
    If multiline Then
        textControl.Height = 44
        textControl.MultiLine = True
        textControl.EnterKeyBehavior = True
        textControl.ScrollBars = fmScrollBarsVertical
    Else
        textControl.Height = 18
    End If
End Sub

Private Sub LoadValues()
    SetText "employee_id", PV("employee_id")
    SetText "event_date", PV("event_date")
    SetText "effective_date", PV("effective_date")
    SetText "order_reference", PV("order_reference")
    SetText "basis_text", PV("basis_text")
    SetText "comment", PV("comment")
    SetText "new_rank", PV("new_rank")
    SetText "new_position", PV("new_position")
    SetText "new_section", PV("new_section")
    SetText "new_military_unit", PV("new_military_unit")
    SetText "new_vus", PV("new_vus")
    SetText "handover_date", PV("handover_date")
    SetText "acceptance_date", PV("acceptance_date")
    SetText "duty_start_date", PV("duty_start_date")
    SetText "destination_unit", PV("destination_unit")
    SetText "destination_location", PV("destination_location")
    mSavedEventID = PV("saved_event_id")
    btnLicenseStatus.Enabled = (mSavedEventID <> "")
    If mSavedEventID <> "" Then SetText "status", t("personnel.wizard.saved_prefix", "Saved:") & " " & mSavedEventID
End Sub

Private Function PV(ByVal fieldKey As String) As String
    Dim rawValue As Variant
    rawValue = mdlPersonnelEvents.GetPersonnelWizardValue(fieldKey)
    If IsDate(rawValue) Then
        PV = Format$(CDate(rawValue), "dd.mm.yyyy")
    Else
        PV = Trim$(CStr(rawValue))
    End If
End Function

Private Sub SetText(ByVal fieldKey As String, ByVal valueText As String)
    Me.Controls("txt_" & fieldKey).Value = valueText
End Sub

Private Function TextOf(ByVal fieldKey As String) As String
    TextOf = Trim$(CStr(Me.Controls("txt_" & fieldKey).Value))
End Function

Private Sub WriteValues()
    Dim fieldKey As Variant
    For Each fieldKey In Array("employee_id", "event_date", "effective_date", "order_reference", "basis_text", "comment", "new_rank", "new_position", "new_section", "new_military_unit", "new_vus", "handover_date", "acceptance_date", "duty_start_date", "destination_unit", "destination_location")
        mdlPersonnelEvents.SetPersonnelWizardValue CStr(fieldKey), TextOf(CStr(fieldKey))
    Next fieldKey
End Sub

Public Function SaveAction() As String
    On Error GoTo Failed
    WriteValues
    mSavedEventID = mdlPersonnelEvents.SavePersonnelWizardAction()
    mdlPersonnelEvents.SetPersonnelWizardValue "saved_event_id", mSavedEventID
    btnLicenseStatus.Enabled = True
    SetText "status", t("personnel.wizard.saved_prefix", "Saved:") & " " & mSavedEventID
    SaveAction = mSavedEventID
    Exit Function
Failed:
    SetText "status", Err.Description
    Application.StatusBar = Err.Description
End Function

Public Function ExportAction() As String
    On Error GoTo Failed
    If mSavedEventID = "" Then Exit Function
    ExportAction = mdlPersonnelEventOrderExport.ExportPersonnelEventOrder(mSavedEventID)
    SetText "status", ExportAction
    Exit Function
Failed:
    SetText "status", Err.Description
    Application.StatusBar = Err.Description
End Function
Private Sub btnExportRequest_Click()
    WriteValues
    If mdlPersonnelEvents.LoadPersonnelWizardCurrentState() Then
        LoadValues
        SetText "status", ""
    Else
        SetText "status", CStr(Application.StatusBar)
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