[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$OutputDirectory,
    [string]$TargetComponentName = 'frmPersonnelActionWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-DesignerLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$Context = @{}
    )

    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = 'New-PersonnelActionWizardV2Designer'
        message = $Message
    }
    foreach ($key in $Context.Keys) { $payload[$key] = $Context[$key] }
    $line = $payload | ConvertTo-Json -Compress -Depth 5
    if ($Level -eq 'DEBUG') { Write-Verbose $line }
    elseif ($Level -eq 'WARN') { Write-Warning $line }
    elseif ($Level -eq 'ERROR') { Write-Error $line }
    else { Write-Host $line }
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Test-VbComponentExists {
    param(
        [Parameter(Mandatory = $true)][object]$VbProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $components = $null
    try {
        $components = $VbProject.VBComponents
        for ($index = 1; $index -le $components.Count; $index++) {
            $component = $null
            try {
                $component = $components.Item($index)
                if ($component.Name -eq $Name) { return $true }
            } finally {
                Release-ComObject $component
            }
        }
    } finally {
        Release-ComObject $components
    }
    return $false
}

if (-not ('PersonnelDesignerNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PersonnelDesignerNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][PersonnelDesignerNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
    return [int]$processId
}

function Stop-OwnedExcelProcessIfNeeded {
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return }

    for ($attempt = 1; $attempt -le 10; $attempt++) {
        if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 250
    }

    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -eq 'EXCEL') {
        Write-DesignerLog WARN 'Excel did not exit after COM Quit; stopping only the process created by this generator.' @{ processId = $ProcessId }
        Stop-Process -Id $ProcessId -Force
    }
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

if (-not $resolvedWorkbook.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook must be inside the CreateOrder project: $resolvedWorkbook"
}
if (-not $resolvedOutput.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Output directory must be inside the CreateOrder project: $resolvedOutput"
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel is running. Close Excel before generating or importing the personnel designer form.'
}

$frmOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frm')
$frxOutput = Join-Path $resolvedOutput ($TargetComponentName + '.frx')
$manifestOutput = Join-Path $resolvedOutput ($TargetComponentName + '.layout.csv')
foreach ($path in @($frmOutput, $frxOutput, $manifestOutput)) {
    if (Test-Path -LiteralPath $path) {
        throw "Target artifact already exists; refusing to overwrite owner layout: $path"
    }
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$buildDirectory = Join-Path $projectRoot ("Trash\personnel-action-designer-v2-build-$stamp")
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\personnel-action-designer-v2-$stamp")
New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null

$buildWorkbook = Join-Path $buildDirectory 'CreateOrder.personnel-designer-build.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-personnel-action-designer-v2.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $buildWorkbook
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook

Write-DesignerLog INFO 'Prepared isolated build workbook and safety backup.' @{
    buildWorkbook = $buildWorkbook
    backupWorkbook = $backupWorkbook
}

$builderCode = @'
Option Explicit

Private Const TARGET_FORM As String = "frmPersonnelActionWizardV2"
Private manifestSheet As Worksheet
Private manifestRow As Long
Private lastStep As String

Public Sub BuildPersonnelActionDesignerV2()
    On Error GoTo ErrorHandler

    Dim targetComponent As Object
    Dim targetDesigner As Object
    Dim wizardFrame As Object
    Dim menuFrame As Object
    Dim multiPage As Object
    Dim transferPage As Object
    Dim exclusionPage As Object

    lastStep = "Prepare manifest"
    Set manifestSheet = PrepareManifestSheet()
    manifestRow = 2

    If ComponentExists(TARGET_FORM) Then Err.Raise 5, , "Target form already exists: " & TARGET_FORM

    lastStep = "Create target form"
    Set targetComponent = ThisWorkbook.VBProject.VBComponents.Add(3)
    targetComponent.Name = TARGET_FORM
    Set targetDesigner = targetComponent.Designer
    targetComponent.Properties.Item("Caption").Value = "Кадровое действие V2"
    targetComponent.Properties.Item("Width").Value = 790
    targetComponent.Properties.Item("Height").Value = 560
    targetComponent.Properties.Item("StartUpPosition").Value = 1

    lastStep = "Create wizard container"
    Set wizardFrame = AddFrame(targetDesigner, "root", "fraWizard", "", 6, 6, 760, 510, True)
    AddSectionLabel wizardFrame, "root/fraWizard", "lbl_section_employee", "1. Найти сотрудника", 12, 10, 710
    AddLabel wizardFrame, "root/fraWizard", "lbl_search", "Фамилия, личный или табельный номер", 12, 32, 220
    AddTextBox wizardFrame, "root/fraWizard", "txt_search", 236, 30, 390, 20, False, False
    AddLabel wizardFrame, "root/fraWizard", "lbl_search_results", "Результат поиска", 12, 58, 110
    AddTextBox wizardFrame, "root/fraWizard", "txt_search_results", 124, 56, 610, 20, False, True
    AddTextBox wizardFrame, "root/fraWizard", "txt_employee_id", 12, 78, 1, 1, False, False, False

    AddSectionLabel wizardFrame, "root/fraWizard", "lbl_section_order", "2. Реквизиты действия", 12, 88, 710
    AddLabel wizardFrame, "root/fraWizard", "lbl_event_date", "Дата события", 12, 110, 95
    AddTextBox wizardFrame, "root/fraWizard", "txt_event_date", 115, 108, 100, 20, False, False
    AddLabel wizardFrame, "root/fraWizard", "lbl_effective_date", "Вступает в силу", 229, 110, 105
    AddTextBox wizardFrame, "root/fraWizard", "txt_effective_date", 342, 108, 100, 20, False, False
    AddLabel wizardFrame, "root/fraWizard", "lbl_order_reference", "Реквизиты приказа", 456, 110, 118
    AddTextBox wizardFrame, "root/fraWizard", "txt_order_reference", 582, 108, 152, 20, False, False
    AddLabel wizardFrame, "root/fraWizard", "lbl_basis_text", "Основание (войдёт в Word)", 12, 138, 170
    AddTextBox wizardFrame, "root/fraWizard", "txt_basis_text", 190, 136, 544, 38, True, False
    AddLabel wizardFrame, "root/fraWizard", "lbl_comment", "Служебный комментарий", 12, 184, 160
    AddTextBox wizardFrame, "root/fraWizard", "txt_comment", 180, 182, 554, 38, True, False

    lastStep = "Create action pages"
    Set multiPage = AddMultiPage(wizardFrame, "root/fraWizard", "mpAction", 12, 226, 722, 204)
    Set transferPage = multiPage.Pages.Item(0)
    transferPage.Name = "pgTransfer"
    transferPage.Caption = "Перемещение"
    RecordPage "root/fraWizard/mpAction", transferPage
    Set exclusionPage = multiPage.Pages.Item(1)
    exclusionPage.Name = "pgExclusion"
    exclusionPage.Caption = "Исключение"
    RecordPage "root/fraWizard/mpAction", exclusionPage

    AddSectionLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_section_transfer", "3. Что меняется при перемещении", 8, 8, 680
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_new_rank", "Новое звание", 8, 30, 100
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_new_rank", 116, 28, 180, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_new_vus", "Новый ВУС", 360, 30, 80
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_new_vus", 448, 28, 180, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_new_position", "Новая должность", 8, 58, 110
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_new_position", 126, 56, 560, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_new_section", "Подразделение", 8, 86, 100
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_new_section", 116, 84, 220, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_new_military_unit", "Воинская часть", 360, 86, 100
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_new_military_unit", 468, 84, 218, 20, False, False
    AddSectionLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_section_dates", "4. Даты и место убытия", 8, 112, 680
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_transfer_handover_date", "Сдал дела", 8, 134, 70
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_transfer_handover_date", 86, 132, 100, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_acceptance_date", "Принял дела", 210, 134, 78
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_acceptance_date", 296, 132, 100, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_duty_start_date", "Приступил", 420, 134, 72
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_duty_start_date", 500, 132, 100, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_transfer_destination_unit", "Куда убывает", 8, 162, 100
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_transfer_destination_unit", 116, 160, 240, 20, False, False
    AddLabel transferPage, "root/fraWizard/mpAction/pgTransfer", "lbl_transfer_destination_location", "Населённый пункт", 376, 162, 105
    AddTextBox transferPage, "root/fraWizard/mpAction/pgTransfer", "txt_transfer_destination_location", 489, 160, 197, 20, False, False

    AddSectionLabel exclusionPage, "root/fraWizard/mpAction/pgExclusion", "lbl_section_exclusion", "3. Сведения об исключении", 8, 8, 680
    AddLabel exclusionPage, "root/fraWizard/mpAction/pgExclusion", "lbl_exclusion_handover_date", "Дата сдачи дел", 8, 30, 95
    AddTextBox exclusionPage, "root/fraWizard/mpAction/pgExclusion", "txt_exclusion_handover_date", 111, 28, 110, 20, False, False
    AddLabel exclusionPage, "root/fraWizard/mpAction/pgExclusion", "lbl_exclusion_destination_unit", "Куда убывает", 244, 30, 100
    AddTextBox exclusionPage, "root/fraWizard/mpAction/pgExclusion", "txt_exclusion_destination_unit", 352, 28, 334, 20, False, False
    AddLabel exclusionPage, "root/fraWizard/mpAction/pgExclusion", "lbl_exclusion_destination_location", "Населённый пункт", 8, 58, 105
    AddTextBox exclusionPage, "root/fraWizard/mpAction/pgExclusion", "txt_exclusion_destination_location", 121, 56, 250, 20, False, False
    AddLabel exclusionPage, "root/fraWizard/mpAction/pgExclusion", "lbl_material_assistance_status", "Материальная помощь за год", 8, 90, 165
    AddTextBox exclusionPage, "root/fraWizard/mpAction/pgExclusion", "txt_material_assistance_status", 181, 88, 505, 20, False, False
    AddLabel exclusionPage, "root/fraWizard/mpAction/pgExclusion", "lbl_main_leave_status", "Основной отпуск за год", 8, 118, 165
    AddTextBox exclusionPage, "root/fraWizard/mpAction/pgExclusion", "txt_main_leave_status", 181, 116, 505, 20, False, False
    AddLabel exclusionPage, "root/fraWizard/mpAction/pgExclusion", "lbl_additional_leave_status", "Дополнительный отпуск за год", 8, 146, 165
    AddTextBox exclusionPage, "root/fraWizard/mpAction/pgExclusion", "txt_additional_leave_status", 181, 144, 505, 20, False, False

    AddTextBox wizardFrame, "root/fraWizard", "txt_status", 12, 438, 722, 20, False, True
    AddButton wizardFrame, "root/fraWizard", "btnExportRequest", "Найти и загрузить", 12, 468, 145, 26
    AddButton wizardFrame, "root/fraWizard", "btnImportResponse", "Сохранить", 169, 468, 120, 26
    AddButton wizardFrame, "root/fraWizard", "btnLicenseStatus", "Экспортировать Word", 301, 468, 180, 26
    AddButton wizardFrame, "root/fraWizard", "btnClose", "Закрыть", 604, 468, 110, 26

    lastStep = "Create menu container"
    Set menuFrame = AddFrame(targetDesigner, "root", "fraActionMenu", "", 140, 80, 490, 350, False)
    AddSectionLabel menuFrame, "root/fraActionMenu", "lblDescription", "Кадровые действия", 24, 20, 420
    AddButton menuFrame, "root/fraActionMenu", "menuEnrollment", "Зачисление", 24, 62, 420, 28
    AddButton menuFrame, "root/fraActionMenu", "menuTransfer", "Перемещение", 24, 104, 420, 28
    AddButton menuFrame, "root/fraActionMenu", "menuExclusion", "Исключение", 24, 146, 420, 28
    AddButton menuFrame, "root/fraActionMenu", "menuHistory", "История сотрудника", 24, 188, 420, 28
    AddButton menuFrame, "root/fraActionMenu", "menuClose", "Закрыть", 316, 240, 128, 28

    manifestSheet.Columns.AutoFit
    Exit Sub

ErrorHandler:
    Dim capturedNumber As Long
    Dim capturedDescription As String
    capturedNumber = Err.Number
    capturedDescription = Err.Description
    On Error Resume Next
    If Not manifestSheet Is Nothing Then
        manifestSheet.Range("M1").Value = capturedNumber
        manifestSheet.Range("N1").Value = capturedDescription
        manifestSheet.Range("O1").Value = lastStep
    End If
    On Error GoTo 0
End Sub

Private Function PrepareManifestSheet() As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("__PersonnelV2Layout").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set PrepareManifestSheet = ThisWorkbook.Worksheets.Add
    PrepareManifestSheet.Name = "__PersonnelV2Layout"
    PrepareManifestSheet.Visible = xlSheetVeryHidden
    PrepareManifestSheet.Range("A1:K1").Value = Array( _
        "container_path", "source_name", "designer_name", "control_type", "caption", _
        "left", "top", "width", "height", "visible", "enabled")
End Function

Private Function AddFrame(ByVal host As Object, ByVal containerPath As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single, ByVal visibleValue As Boolean) As Object
    Set AddFrame = host.Controls.Add("Forms.Frame.1", controlName, True)
    AddFrame.Caption = captionText
    AddFrame.Left = leftValue: AddFrame.Top = topValue
    AddFrame.Width = widthValue: AddFrame.Height = heightValue
    AddFrame.Visible = visibleValue
    RecordControl containerPath, AddFrame, "Frame"
End Function

Private Function AddMultiPage(ByVal host As Object, ByVal containerPath As String, ByVal controlName As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single) As Object
    Set AddMultiPage = host.Controls.Add("Forms.MultiPage.1", controlName, True)
    AddMultiPage.Left = leftValue: AddMultiPage.Top = topValue
    AddMultiPage.Width = widthValue: AddMultiPage.Height = heightValue
    RecordControl containerPath, AddMultiPage, "MultiPage"
End Function

Private Sub AddLabel(ByVal host As Object, ByVal containerPath As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single)
    Dim controlItem As Object
    Set controlItem = host.Controls.Add("Forms.Label.1", controlName, True)
    controlItem.Caption = captionText
    controlItem.Left = leftValue: controlItem.Top = topValue
    controlItem.Width = widthValue: controlItem.Height = 18
    RecordControl containerPath, controlItem, "Label"
End Sub

Private Sub AddSectionLabel(ByVal host As Object, ByVal containerPath As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single)
    Dim controlItem As Object
    Set controlItem = host.Controls.Add("Forms.Label.1", controlName, True)
    controlItem.Caption = captionText
    controlItem.Left = leftValue: controlItem.Top = topValue
    controlItem.Width = widthValue: controlItem.Height = 18
    controlItem.Font.Bold = True: controlItem.Font.Size = 10
    RecordControl containerPath, controlItem, "Label"
End Sub

Private Sub AddTextBox(ByVal host As Object, ByVal containerPath As String, ByVal controlName As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single, ByVal multilineValue As Boolean, ByVal lockedValue As Boolean, Optional ByVal visibleValue As Boolean = True)
    Dim controlItem As Object
    Set controlItem = host.Controls.Add("Forms.TextBox.1", controlName, True)
    controlItem.Left = leftValue: controlItem.Top = topValue
    controlItem.Width = widthValue: controlItem.Height = heightValue
    controlItem.MultiLine = multilineValue
    controlItem.Locked = lockedValue
    controlItem.Visible = visibleValue
    If multilineValue Then
        controlItem.EnterKeyBehavior = True
        controlItem.ScrollBars = 2
    End If
    If lockedValue Then controlItem.BackColor = RGB(242, 242, 242)
    RecordControl containerPath, controlItem, "TextBox"
End Sub

Private Sub AddButton(ByVal host As Object, ByVal containerPath As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single)
    Dim controlItem As Object
    Set controlItem = host.Controls.Add("Forms.CommandButton.1", controlName, True)
    controlItem.Caption = captionText
    controlItem.Left = leftValue: controlItem.Top = topValue
    controlItem.Width = widthValue: controlItem.Height = heightValue
    RecordControl containerPath, controlItem, "CommandButton"
End Sub

Private Sub RecordPage(ByVal containerPath As String, ByVal pageItem As Object)
    manifestSheet.Cells(manifestRow, 1).Value = containerPath
    manifestSheet.Cells(manifestRow, 2).Value = pageItem.Name
    manifestSheet.Cells(manifestRow, 3).Value = pageItem.Name
    manifestSheet.Cells(manifestRow, 4).Value = "Page"
    manifestSheet.Cells(manifestRow, 5).Value = pageItem.Caption
    manifestSheet.Cells(manifestRow, 10).Value = True
    manifestSheet.Cells(manifestRow, 11).Value = True
    manifestRow = manifestRow + 1
End Sub

Private Sub RecordControl(ByVal containerPath As String, ByVal controlItem As Object, ByVal controlType As String)
    manifestSheet.Cells(manifestRow, 1).Value = containerPath
    manifestSheet.Cells(manifestRow, 2).Value = controlItem.Name
    manifestSheet.Cells(manifestRow, 3).Value = controlItem.Name
    manifestSheet.Cells(manifestRow, 4).Value = controlType
    On Error Resume Next
    manifestSheet.Cells(manifestRow, 5).Value = controlItem.Caption
    On Error GoTo 0
    manifestSheet.Cells(manifestRow, 6).Value = controlItem.Left
    manifestSheet.Cells(manifestRow, 7).Value = controlItem.Top
    manifestSheet.Cells(manifestRow, 8).Value = controlItem.Width
    manifestSheet.Cells(manifestRow, 9).Value = controlItem.Height
    manifestSheet.Cells(manifestRow, 10).Value = controlItem.Visible
    manifestSheet.Cells(manifestRow, 11).Value = controlItem.Enabled
    manifestRow = manifestRow + 1
End Sub

Private Function ComponentExists(ByVal componentName As String) As Boolean
    Dim component As Object
    For Each component In ThisWorkbook.VBProject.VBComponents
        If StrComp(component.Name, componentName, vbTextCompare) = 0 Then
            ComponentExists = True
            Exit Function
        End If
    Next component
End Function
'@

$formCode = @'
Option Explicit

Private Const DEBUG_LOGGING As Boolean = True

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
    failureNumber = Err.Number
    failureDescription = Err.Description
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
        multiPage.Value = 1
    Else
        Me.Caption = t("personnel.wizard.title.transfer", "Кадровое перемещение")
        multiPage.Value = 0
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
    Set foundControl = containerHost.Controls.Item(controlName)
    On Error GoTo 0
    If Not foundControl Is Nothing Then
        Set FindDesignerControlInContainer = foundControl
        Exit Function
    End If

    For Each controlItem In containerHost.Controls
        If TypeName(controlItem) = "MultiPage" Then
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
        ElseIf TypeName(controlItem) = "Frame" Then
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
    targetControl.Value = valueText
End Sub

Private Function TextOf(ByVal fieldKey As String) As String
    TextOf = Trim$(CStr(FindDesignerControl(DesignerNameForField(fieldKey)).Value))
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
    LogDebug "save-error", "number=" & CStr(Err.Number) & "; description=" & Err.Description
    SetText "status", Err.Description
    Application.StatusBar = Err.Description
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
    LogDebug "export-error", "event_id=" & mSavedEventID & "; number=" & CStr(Err.Number) & "; description=" & Err.Description
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

Private Sub LogDebug(ByVal actionName As String, ByVal detailText As String)
    If DEBUG_LOGGING Then Debug.Print "[PERSONNEL-ACTION-V2] action=" & actionName & "; " & detailText
End Sub
'@

$excel = $null
$excelProcessId = 0
$buildBook = $null
$builderComponent = $null
$targetComponent = $null
$manifestSheet = $null
$usedRange = $null
$currentStep = 'start isolated generation'
try {
    Write-DesignerLog INFO 'Opening isolated workbook and generating design-time form.' @{ workbook = $buildWorkbook }
    $currentStep = 'create Excel application'
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $currentStep = 'open isolated workbook'
    $buildBook = $excel.Workbooks.Open($buildWorkbook, 0, $false)

    if ($buildBook.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot create the designer form.' }
    if (Test-VbComponentExists -VbProject $buildBook.VBProject -Name $TargetComponentName) {
        throw "Target component already exists in the build workbook: $TargetComponentName"
    }

    $currentStep = 'add builder module'
    $builderComponent = $buildBook.VBProject.VBComponents.Add(1)
    $builderComponent.Name = 'modPersonnelV2Builder'
    $currentStep = 'write builder module'
    $builderComponent.CodeModule.AddFromString($builderCode)
    $currentStep = 'run builder module'
    $excel.Run("'$($buildBook.Name)'!modPersonnelV2Builder.BuildPersonnelActionDesignerV2")

    $currentStep = 'read builder manifest'
    $manifestSheet = $buildBook.Worksheets.Item('__PersonnelV2Layout')
    $builderErrorNumber = [int]$manifestSheet.Range('M1').Value2
    if ($builderErrorNumber -ne 0) {
        $builderErrorDescription = [string]$manifestSheet.Range('N1').Value2
        $builderErrorStep = [string]$manifestSheet.Range('O1').Value2
        throw "VBA builder failed at '$builderErrorStep': [$builderErrorNumber] $builderErrorDescription"
    }

    $currentStep = 'load generated component'
    $targetComponent = $buildBook.VBProject.VBComponents.Item($TargetComponentName)
    if ($targetComponent.Type -ne 3) { throw "Generated component has unexpected type: $($targetComponent.Type)" }
    $currentStep = 'write V2 form code'
    $targetComponent.CodeModule.AddFromString($formCode)
    $currentStep = 'export V2 form'
    $targetComponent.Export($frmOutput)
    if (-not (Test-Path -LiteralPath $frmOutput) -or -not (Test-Path -LiteralPath $frxOutput)) {
        throw 'Excel did not export the expected .frm/.frx pair.'
    }

    $currentStep = 'export layout manifest'
    $usedRange = $manifestSheet.UsedRange
    $values = $usedRange.Value2
    $rows = @()
    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
        $rows += [pscustomobject]@{
            container_path = [string]$values[$row, 1]
            source_name = [string]$values[$row, 2]
            designer_name = [string]$values[$row, 3]
            control_type = [string]$values[$row, 4]
            caption = [string]$values[$row, 5]
            left = $values[$row, 6]
            top = $values[$row, 7]
            width = $values[$row, 8]
            height = $values[$row, 9]
            visible = $values[$row, 10]
            enabled = $values[$row, 11]
        }
    }
    $rows | Export-Csv -LiteralPath $manifestOutput -NoTypeInformation -Encoding utf8
    $currentStep = 'save isolated workbook'
    $buildBook.Save()

    Write-DesignerLog INFO 'Generated and exported personnel design-time form.' @{
        controlsAndPages = $rows.Count
        frm = $frmOutput
        frx = $frxOutput
        manifest = $manifestOutput
    }
} catch {
    Write-DesignerLog ERROR 'Personnel designer generation failed.' @{ step = $currentStep; error = $_.Exception.Message }
    throw
} finally {
    if ($buildBook) { $buildBook.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $usedRange
    Release-ComObject $manifestSheet
    Release-ComObject $targetComponent
    Release-ComObject $builderComponent
    Release-ComObject $buildBook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}

for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel remained running after isolated generation; refusing to import into the working workbook.'
}

& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardV2Designer.ps1') -WorkbookPath $buildWorkbook -SourceDirectory $resolvedOutput -TargetComponentName $TargetComponentName

$excel = $null
$excelProcessId = 0
$workingBook = $null
$importedComponent = $null
try {
    Write-DesignerLog INFO 'Importing verified personnel designer form into working workbook.' @{ workbook = $resolvedWorkbook }
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workingBook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)

    if ($workingBook.VBProject.Protection -ne 0) { throw 'Working VBA project is protected; cannot import the designer form.' }
    if (Test-VbComponentExists -VbProject $workingBook.VBProject -Name $TargetComponentName) {
        throw "Target component already exists in the working workbook: $TargetComponentName"
    }

    $importedComponent = $workingBook.VBProject.VBComponents.Import($frmOutput)
    if ($importedComponent.Name -ne $TargetComponentName -or $importedComponent.Type -ne 3) {
        throw "Imported component verification failed: name=$($importedComponent.Name), type=$($importedComponent.Type)"
    }
    if (-not (Test-VbComponentExists -VbProject $workingBook.VBProject -Name 'frmPersonnelActionWizard')) {
        throw 'Original frmPersonnelActionWizard is missing after V2 import.'
    }

    $workingBook.Save()
    Write-DesignerLog INFO 'Personnel designer V2 installed; original form remains active.' @{
        workbook = $resolvedWorkbook
        component = $TargetComponentName
        backup = $backupWorkbook
    }
} catch {
    Write-DesignerLog ERROR 'Personnel designer form import failed.' @{ error = $_.Exception.Message; backup = $backupWorkbook }
    throw
} finally {
    if ($workingBook) { $workingBook.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $importedComponent
    Release-ComObject $workingBook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}

for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel remained running after personnel V2 import.'
}

& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardV2Designer.ps1') -WorkbookPath $resolvedWorkbook -SourceDirectory $resolvedOutput -TargetComponentName $TargetComponentName

Write-DesignerLog INFO 'Personnel action designer V2 generation completed.' @{
    component = $TargetComponentName
    backup = $backupWorkbook
    buildWorkbook = $buildWorkbook
}
