[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmPersonnelActionWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-VbaText {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $true)][string]$ModulePath
    )
    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($code)
}

function Import-UserForm {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$FormName,
        [Parameter(Mandatory = $true)][string]$FormPath
    )
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item($FormName)) } catch {}
    $component = $Workbook.VBProject.VBComponents.Import($FormPath)
    if ($component.Name -ne $FormName -or $component.Type -ne 3) {
        throw "Imported form verification failed: name=$($component.Name), type=$($component.Type)"
    }
}

function Get-DocxText {
    param([Parameter(Mandatory = $true)][string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry('word/document.xml')
        if ($null -eq $entry) { throw "DOCX does not contain word/document.xml: $Path" }
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
        try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $xmlText = $xmlText -replace '</w:p>', "`n" -replace '<w:tab[^>]*/>', "`t"
        return [Net.WebUtility]::HtmlDecode(($xmlText -replace '<[^>]+>', ''))
    } finally {
        $archive.Dispose()
    }
}

function Assert-Contains {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Expected,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if ($Text -notlike "*$Expected*") { throw $Message }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
    throw 'Excel or Word is running. Close Office applications before the personnel V2 E2E test.'
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$formPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
if (-not (Test-Path -LiteralPath $formPath)) { throw "Missing personnel V2 form source: $formPath" }

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testDirectory = Join-Path $projectRoot ("Trash\personnel-action-v2-e2e-$stamp")
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$testWorkbookPath = Join-Path $testDirectory 'CreateOrder.personnel-action-v2-e2e.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $testWorkbookPath

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)

    Import-CodeModuleText -Workbook $workbook -ModuleName 'ModuleLocalization' -ModulePath (Join-Path $resolvedSource 'ModuleLocalization.bas')
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelEvents' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelEvents.bas')
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelEventOrderExport' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelEventOrderExport.bas')
    Import-UserForm -Workbook $workbook -FormName $TargetComponentName -FormPath $formPath

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item('personnel_v2_e2e_probe')) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = 'personnel_v2_e2e_probe'
    $probeCode = @'
Option Explicit

Private Function FindProbeControl(ByVal containerHost As Object, ByVal controlName As String) As Object
    Dim foundControl As Object
    Dim controlItem As Object
    Dim pageItem As Object

    On Error Resume Next
    Set foundControl = containerHost.Controls.Item(controlName)
    On Error GoTo 0
    If Not foundControl Is Nothing Then
        Set FindProbeControl = foundControl
        Exit Function
    End If

    For Each controlItem In containerHost.Controls
        If TypeName(controlItem) = "MultiPage" Then
            For Each pageItem In controlItem.Pages
                Set foundControl = FindProbeControl(pageItem, controlName)
                If Not foundControl Is Nothing Then
                    Set FindProbeControl = foundControl
                    Exit Function
                End If
            Next pageItem
        ElseIf TypeName(controlItem) = "Frame" Then
            Set foundControl = FindProbeControl(controlItem, controlName)
            If Not foundControl Is Nothing Then
                Set FindProbeControl = foundControl
                Exit Function
            End If
        End If
    Next controlItem
End Function

Private Sub SetProbeValue(ByVal formObject As Object, ByVal controlName As String, ByVal valueText As String)
    Dim targetControl As Object
    Set targetControl = FindProbeControl(formObject, controlName)
    If targetControl Is Nothing Then Err.Raise 930, , "Missing V2 control: " & controlName
    targetControl.Value = valueText
End Sub

Public Function RunPersonnelActionV2E2E() As String
    Dim formObject As Object
    Dim enrollmentID As String
    Dim transferID As String
    Dim exclusionID As String
    Dim employeeID As String
    Dim transferPath As String
    Dim exclusionPath As String
    Dim currentState As Object
    Dim employeeRow As Long

    On Error GoTo Failed
    mdlPersonnelEvents.ResetPersonnelEventInput
    mdlPersonnelEvents.SetPersonnelWizardValue "event_type", "ENROLLMENT"
    mdlPersonnelEvents.SetPersonnelWizardValue "event_date", DateSerial(2026, 8, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "effective_date", DateSerial(2026, 8, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "order_reference", "V2-ENROLL-001"
    mdlPersonnelEvents.SetPersonnelWizardValue "basis_text", "Personnel V2 E2E enrollment"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_fio", "Personnel V2 Test Employee"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_personal_number", "V2-001"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_rank", "Private"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_position", "Initial V2 position"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_section", "Initial V2 section"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_military_unit", "V2 unit"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_vus", "100100"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_tariff_rank", "5"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_position_salary", "25000"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_rank_salary", "10000"
    enrollmentID = mdlPersonnelEvents.SavePersonnelEventInput(False)
    employeeID = CStr(mdlPersonnelEvents.GetPersonnelWizardValue("employee_id"))
    If enrollmentID = "" Or employeeID = "" Then Err.Raise 931, , "Enrollment fixture was not created"

    mdlPersonnelEvents.PrepareNewPersonnelAction "TRANSFER"
    Load __TARGET__
    Set formObject = __TARGET__
    SetProbeValue formObject, "txt_search", "Personnel V2 Test"
    If InStr(1, CStr(FindProbeControl(formObject, "txt_search_results").Value), "Personnel V2 Test", vbTextCompare) = 0 Then Err.Raise 932, , "V2 search preview did not find the employee"
    If formObject.ExportAction() <> "" Then Err.Raise 933, , "V2 export guard failed before save"
    SetProbeValue formObject, "txt_employee_id", employeeID
    SetProbeValue formObject, "txt_event_date", "02.08.2026"
    SetProbeValue formObject, "txt_effective_date", "03.08.2026"
    SetProbeValue formObject, "txt_order_reference", "V2-TRANSFER-001"
    SetProbeValue formObject, "txt_basis_text", "Personnel V2 E2E transfer"
    SetProbeValue formObject, "txt_new_rank", "Private"
    SetProbeValue formObject, "txt_new_position", "Transferred V2 position"
    SetProbeValue formObject, "txt_new_section", "Transferred V2 section"
    SetProbeValue formObject, "txt_new_military_unit", "V2 unit 2"
    SetProbeValue formObject, "txt_new_vus", "200200"
    SetProbeValue formObject, "txt_transfer_handover_date", "02.08.2026"
    SetProbeValue formObject, "txt_acceptance_date", "03.08.2026"
    SetProbeValue formObject, "txt_duty_start_date", "04.08.2026"
    SetProbeValue formObject, "txt_transfer_destination_unit", "V2 destination unit"
    SetProbeValue formObject, "txt_transfer_destination_location", "V2 destination city"
    transferID = formObject.SaveAction()
    If transferID = "" Then Err.Raise 934, , "V2 transfer did not return EventID"
    Set currentState = mdlPersonnelEvents.GetCurrentPersonnelState(employeeID)
    If CStr(currentState("position")) <> "Transferred V2 position" Then Err.Raise 935, , "V2 transfer did not update current state"
    transferPath = formObject.ExportAction()
    If transferPath = "" Then Err.Raise 936, , "V2 transfer did not export Word"
    Unload formObject

    mdlPersonnelEvents.PrepareNewPersonnelAction "EXCLUSION"
    Load __TARGET__
    Set formObject = __TARGET__
    SetProbeValue formObject, "txt_employee_id", employeeID
    SetProbeValue formObject, "txt_event_date", "05.08.2026"
    SetProbeValue formObject, "txt_effective_date", "06.08.2026"
    SetProbeValue formObject, "txt_order_reference", "V2-EXCLUSION-001"
    SetProbeValue formObject, "txt_basis_text", "Personnel V2 E2E exclusion"
    SetProbeValue formObject, "txt_exclusion_handover_date", "05.08.2026"
    SetProbeValue formObject, "txt_exclusion_destination_unit", "V2 destination unit"
    SetProbeValue formObject, "txt_exclusion_destination_location", "V2 destination city"
    SetProbeValue formObject, "txt_material_assistance_status", "оказана"
    SetProbeValue formObject, "txt_main_leave_status", "использован"
    SetProbeValue formObject, "txt_additional_leave_status", "не использован"
    exclusionID = formObject.SaveAction()
    If exclusionID = "" Then Err.Raise 937, , "V2 exclusion did not return EventID"
    For employeeRow = 2 To ThisWorkbook.Worksheets("Employees").Cells(ThisWorkbook.Worksheets("Employees").Rows.Count, 1).End(xlUp).Row
        If CStr(ThisWorkbook.Worksheets("Employees").Cells(employeeRow, 1).Value) = employeeID Then Exit For
    Next employeeRow
    If CStr(ThisWorkbook.Worksheets("Employees").Cells(employeeRow, 10).Value) <> "NO" Then Err.Raise 938, , "V2 exclusion did not deactivate employee"
    exclusionPath = formObject.ExportAction()
    If exclusionPath = "" Then Err.Raise 939, , "V2 exclusion did not export Word"
    Unload formObject

    RunPersonnelActionV2E2E = "PERSONNEL_ACTION_V2_E2E_OK|" & transferPath & "|" & exclusionPath
    Exit Function
Failed:
    RunPersonnelActionV2E2E = "FAILED: " & Err.Description
End Function
'@.Replace('__TARGET__', $TargetComponentName)
    $probe.CodeModule.AddFromString($probeCode)

    $result = [string]$excel.Run("'$($workbook.Name)'!personnel_v2_e2e_probe.RunPersonnelActionV2E2E")
    if ($result -notlike 'PERSONNEL_ACTION_V2_E2E_OK|*') { throw $result }
    $parts = $result -split '\|', 3
    if ($parts.Count -ne 3) { throw "Unexpected V2 E2E result: $result" }
    $transferPath = $parts[1]
    $exclusionPath = $parts[2]
    if (-not (Test-Path -LiteralPath $transferPath)) { throw "V2 transfer DOCX missing: $transferPath" }
    if (-not (Test-Path -LiteralPath $exclusionPath)) { throw "V2 exclusion DOCX missing: $exclusionPath" }

    $transferText = Get-DocxText -Path $transferPath
    $exclusionText = Get-DocxText -Path $exclusionPath
    Assert-Contains $transferText '02.08.2026 г. № V2-TRANSFER-001' 'V2 transfer DOCX omitted the actual order date/number.'
    Assert-Contains $transferText 'Transferred V2 position' 'V2 transfer DOCX omitted the new position.'
    Assert-Contains $transferText 'ВУС-200200' 'V2 transfer DOCX omitted the new VUS.'
    Assert-Contains $exclusionText '05.08.2026 г. № V2-EXCLUSION-001' 'V2 exclusion DOCX omitted the actual order date/number.'
    Assert-Contains $exclusionText 'Материальная помощь за текущий год: оказана' 'V2 exclusion DOCX omitted material assistance status.'
    Assert-Contains $exclusionText 'Дополнительный отпуск за текущий год: не использован' 'V2 exclusion DOCX omitted additional leave status.'

    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    Write-Output "PERSONNEL_ACTION_V2_E2E_OK|$transferPath|$exclusionPath"
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
    throw 'Office process remained running after the personnel V2 E2E test.'
}
