[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmPersonnelActionWizardV2',
    [ValidateSet('V1', 'V2')][string]$ExpectedActiveVersion = 'V1'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Condition {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) { throw $Message }
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

if (-not ('PersonnelDesignerTestNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PersonnelDesignerTestNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][PersonnelDesignerTestNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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
    if ($process -and $process.ProcessName -eq 'EXCEL') { Stop-Process -Id $ProcessId -Force }
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$frmPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$frxPath = Join-Path $resolvedSource ($TargetComponentName + '.frx')
$manifestPath = Join-Path $resolvedSource ($TargetComponentName + '.layout.csv')

foreach ($path in @($frmPath, $frxPath, $manifestPath)) {
    Assert-Condition (Test-Path -LiteralPath $path) "Missing personnel V2 artifact: $path"
}
Assert-Condition (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) 'Excel is running. Close Excel before the personnel V2 designer test.'

$bytes = [IO.File]::ReadAllBytes($frmPath)
$formText = [Text.Encoding]::GetEncoding(1251).GetString($bytes)
Assert-Condition ($formText.Contains(('Attribute VB_Name = "{0}"' -f $TargetComponentName))) 'The exported .frm has an unexpected VB_Name.'
Assert-Condition ($formText.Contains('Private Sub BindDesignerControls()')) 'The V2 form is missing design-time control binding.'
Assert-Condition ($formText.Contains('Private Sub ApplyDesignerLocalization()')) 'The V2 form is missing designer localization.'
Assert-Condition (-not $formText.Contains('Controls.Add(')) 'The V2 form code must not create runtime controls.'
Assert-Condition (-not [regex]::IsMatch($formText, '\.(Left|Top|Width|Height)\s*=')) 'The V2 form code must not override owner geometry.'
Assert-Condition (-not [regex]::IsMatch($formText, '(?<!\r)\n')) 'The V2 .frm must use CRLF line endings.'

$manifest = @(Import-Csv -LiteralPath $manifestPath)
$pages = @($manifest | Where-Object control_type -eq 'Page')
$designerNames = @($manifest | ForEach-Object designer_name)
$uniqueDesignerNames = @($designerNames | Sort-Object -Unique)
Assert-Condition ($manifest.Count -ge 50) "Layout manifest is unexpectedly small: $($manifest.Count) rows."
Assert-Condition ($pages.Count -eq 2) "Expected two action pages in the manifest; found $($pages.Count)."
Assert-Condition ($designerNames.Count -eq $uniqueDesignerNames.Count) 'Designer control names must be globally unique.'

$requiredNames = @(
    'fraWizard', 'fraActionMenu', 'mpAction', 'pgTransfer', 'pgExclusion',
    'txt_search', 'txt_search_results', 'txt_employee_id', 'txt_event_date',
    'txt_effective_date', 'txt_order_reference', 'txt_basis_text', 'txt_comment',
    'txt_new_rank', 'txt_new_position', 'txt_new_section', 'txt_new_military_unit',
    'txt_new_vus', 'txt_transfer_handover_date', 'txt_acceptance_date',
    'txt_duty_start_date', 'txt_transfer_destination_unit',
    'txt_transfer_destination_location', 'txt_exclusion_handover_date',
    'txt_exclusion_destination_unit', 'txt_exclusion_destination_location',
    'txt_material_assistance_status', 'txt_main_leave_status',
    'txt_additional_leave_status', 'txt_status', 'btnExportRequest',
    'btnImportResponse', 'btnLicenseStatus', 'btnClose', 'menuEnrollment',
    'menuTransfer', 'menuExclusion', 'menuHistory', 'menuClose'
)
foreach ($requiredName in $requiredNames) {
    Assert-Condition ($requiredName -in $designerNames) "Missing design-time control: $requiredName"
}

$excel = $null
$excelProcessId = 0
$book = $null
$v1 = $null
$v2 = $null
$designer = $null
$wizardFrame = $null
$multiPage = $null
$probeComponent = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)

    $v1 = $book.VBProject.VBComponents.Item('frmPersonnelActionWizard')
    $v2 = $book.VBProject.VBComponents.Item($TargetComponentName)
    Assert-Condition ($v1.Type -eq 3) "Original form has unexpected component type: $($v1.Type)"
    Assert-Condition ($v2.Type -eq 3) "V2 has unexpected component type: $($v2.Type)"
    $designer = $v2.Designer
    Assert-Condition ([double]$designer.InsideWidth -ge 740) "V2 form client width is unexpectedly small: $($designer.InsideWidth)"
    Assert-Condition ([double]$designer.InsideHeight -ge 500) "V2 form client height is unexpectedly small: $($designer.InsideHeight)"

    $wizardFrame = $designer.Controls.Item('fraWizard')
    $multiPage = $wizardFrame.Controls.Item('mpAction')
    Assert-Condition ($multiPage.Pages.Count -eq 2) "Installed V2 must contain two pages; found $($multiPage.Pages.Count)."
    Assert-Condition ($multiPage.Pages.Item(0).Name -eq 'pgTransfer') 'The first V2 page must be pgTransfer.'
    Assert-Condition ($multiPage.Pages.Item(1).Name -eq 'pgExclusion') 'The second V2 page must be pgExclusion.'

    $v2Code = $v2.CodeModule.Lines(1, $v2.CodeModule.CountOfLines)
    Assert-Condition ($v2Code.Contains('Private Sub BindDesignerControls()')) 'Installed V2 is not connected to design-time controls.'
    Assert-Condition (-not $v2Code.Contains('Controls.Add(')) 'Installed V2 unexpectedly contains runtime control creation.'

    $eventsComponent = $book.VBProject.VBComponents.Item('mdlPersonnelEvents')
    try {
        $eventsCode = $eventsComponent.CodeModule.Lines(1, $eventsComponent.CodeModule.CountOfLines)
        if ($ExpectedActiveVersion -eq 'V2') {
            Assert-Condition ($eventsCode.Contains(($TargetComponentName + '.Show'))) 'The active personnel route does not open V2.'
            Assert-Condition (-not $eventsCode.Contains('frmPersonnelActionWizard.Show')) 'The active personnel route still opens V1.'
        } else {
            Assert-Condition ($eventsCode.Contains('frmPersonnelActionWizard.Show')) 'The current V1 route is missing before owner acceptance.'
            Assert-Condition (-not $eventsCode.Contains(($TargetComponentName + '.Show'))) 'V2 must not become active before owner layout acceptance.'
        }
    } finally {
        Release-ComObject $eventsComponent
    }

    $probeCode = @'
Option Explicit

Public Function RunPersonnelActionV2DesignerProbe() As String
    Dim formObject As Object
    Dim wizardFrame As Object
    Dim menuFrame As Object
    Dim multiPage As Object

    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure

    mdlPersonnelEvents.PreparePersonnelActionMenu
    Load __TARGET__
    Set formObject = __TARGET__
    If Not formObject.IsActionMenu Then Err.Raise 901, , "V2 did not enter action-menu mode"
    Set menuFrame = formObject.Controls.Item("fraActionMenu")
    If Not menuFrame.Visible Then Err.Raise 902, , "Action-menu frame is not visible"
    Unload formObject

    mdlPersonnelEvents.PrepareNewPersonnelAction "TRANSFER"
    Load __TARGET__
    Set formObject = __TARGET__
    If formObject.IsActionMenu Then Err.Raise 903, , "TRANSFER unexpectedly entered menu mode"
    Set wizardFrame = formObject.Controls.Item("fraWizard")
    Set multiPage = wizardFrame.Controls.Item("mpAction")
    If multiPage.Value <> 0 Then Err.Raise 904, , "TRANSFER page was not selected"
    If wizardFrame.Controls.Item("btnLicenseStatus").Enabled Then Err.Raise 905, , "Export must stay disabled before save"
    Unload formObject

    mdlPersonnelEvents.PrepareNewPersonnelAction "EXCLUSION"
    Load __TARGET__
    Set formObject = __TARGET__
    Set wizardFrame = formObject.Controls.Item("fraWizard")
    Set multiPage = wizardFrame.Controls.Item("mpAction")
    If multiPage.Value <> 1 Then Err.Raise 906, , "EXCLUSION page was not selected"
    Unload formObject

    RunPersonnelActionV2DesignerProbe = "PERSONNEL_ACTION_V2_DESIGNER_OK"
End Function
'@.Replace('__TARGET__', $TargetComponentName)

    $probeComponent = $book.VBProject.VBComponents.Add(1)
    $probeComponent.Name = 'modPersonnelV2Probe'
    $probeComponent.CodeModule.AddFromString($probeCode)
    $probeResult = [string]$excel.Run("'$($book.Name)'!modPersonnelV2Probe.RunPersonnelActionV2DesignerProbe")
    Assert-Condition ($probeResult -eq 'PERSONNEL_ACTION_V2_DESIGNER_OK') "Unexpected V2 probe result: $probeResult"
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $probeComponent
    Release-ComObject $multiPage
    Release-ComObject $wizardFrame
    Release-ComObject $designer
    Release-ComObject $v2
    Release-ComObject $v1
    Release-ComObject $book
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}

Assert-Condition (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) 'Excel remained running after personnel V2 designer verification.'
Write-Host ("Personnel Action Wizard V2 designer verification passed: {0} manifest rows, 2 pages, V1 retained, {1} active." -f $manifest.Count, $ExpectedActiveVersion)
