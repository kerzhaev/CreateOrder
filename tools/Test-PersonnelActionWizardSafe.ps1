param(
    [string]$WorkbookPath,
    [switch]$RequireInstalledLocalization
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }

$ErrorActionPreference = "Stop"
$workspace = Split-Path -Parent $PSScriptRoot
$testDirectory = Join-Path $workspace "_tmp_personnel_action_wizard_test"
$testWorkbookPath = Join-Path $testDirectory "CreateOrder_personnel_action_wizard_test.xlsm"
$moduleDirectory = Join-Path $workspace "CreateOrder.xlsm.modules"

function Read-VbaText([string]$Path) {
    [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText($Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = Read-VbaText $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($code)
}

function Import-UserForm($Workbook, [string]$FormName, [string]$FormPath) {
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item($FormName)) } catch {}
    $component = $Workbook.VBProject.VBComponents.Import($FormPath)
    if ($component.Name -ne $FormName) { throw "Imported form name mismatch: $($component.Name)" }
}

function Get-DocxText([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry("word/document.xml")
        if ($null -eq $entry) { throw "DOCX does not contain word/document.xml: $Path" }
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
        try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $xmlText = $xmlText -replace '</w:p>', "`n" -replace '<w:tab[^>]*/>', "`t"
        return [Net.WebUtility]::HtmlDecode(($xmlText -replace '<[^>]+>', ''))
    }
    finally {
        $archive.Dispose()
    }
}

function Assert-DocxContains([string]$Text, [string]$Expected, [string]$Message) {
    if ($Text -notlike "*$Expected*") { throw $Message }
}

function Assert-WorkbookLocalizationEntry($Workbook, [string]$Key) {
    $sheet = $null
    try {
        $sheet = $Workbook.Worksheets.Item('Localization')
        $lastRow = [int]$sheet.Cells($sheet.Rows.Count, 1).End(-4162).Row
        for ($row = 2; $row -le $lastRow; $row++) {
            if ([string]::Equals(([string]$sheet.Cells($row, 1).Value2).Trim(), $Key, [StringComparison]::OrdinalIgnoreCase)) {
                if ([string]::IsNullOrWhiteSpace([string]$sheet.Cells($row, 2).Value2)) { throw "Localization value is blank: $Key" }
                return
            }
        }
        throw "Localization key is missing from the workbook sheet: $Key"
    }
    finally {
        if ($null -ne $sheet -and [Runtime.InteropServices.Marshal]::IsComObject($sheet)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
        }
    }
}

New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
Copy-Item -LiteralPath $WorkbookPath -Destination $testWorkbookPath -Force
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 1 } catch {}
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)
    if ($RequireInstalledLocalization) {
        Assert-WorkbookLocalizationEntry $workbook 'personnel.wizard.material_assistance_status'
        Assert-WorkbookLocalizationEntry $workbook 'personnel.wizard.main_leave_status'
        Assert-WorkbookLocalizationEntry $workbook 'personnel.wizard.additional_leave_status'
    }
    Import-CodeModuleText $workbook "ModuleLocalization" (Join-Path $moduleDirectory "ModuleLocalization.bas")
    Import-CodeModuleText $workbook "mdlPersonnelEvents" (Join-Path $moduleDirectory "mdlPersonnelEvents.bas")
    Import-CodeModuleText $workbook "mdlPersonnelEventOrderExport" (Join-Path $moduleDirectory "mdlPersonnelEventOrderExport.bas")
    Import-UserForm $workbook "frmPersonnelActionWizard" (Join-Path $moduleDirectory "frmPersonnelActionWizard.frm")

    try { $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item("personnel_action_wizard_probe")) } catch {}
    $probe = $workbook.VBProject.VBComponents.Add(1)
    $probe.Name = "personnel_action_wizard_probe"
    $probe.CodeModule.AddFromString(@"
Option Explicit
Public Function ProbePersonnelActionWizard() As String
    Dim enrollmentID As String, duplicateEnrollmentID As String, transferID As String, exclusionID As String, outputPath As String, transferPath As String, exclusionPath As String, employeeID As String, currentState As Object, employeeRow As Long, searchMatches As Collection
    On Error GoTo Failed
    mdlPersonnelEvents.PreparePersonnelActionMenu
    Load frmPersonnelActionWizard
    If Not frmPersonnelActionWizard.IsActionMenu Then Err.Raise 698, , "Personnel action menu did not initialize in selection mode"
    Unload frmPersonnelActionWizard
    mdlPersonnelEvents.ResetPersonnelEventInput
    mdlPersonnelEvents.SetPersonnelWizardValue "event_type", "ENROLLMENT"
    mdlPersonnelEvents.SetPersonnelWizardValue "event_date", DateSerial(2026, 7, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "effective_date", DateSerial(2026, 7, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "order_reference", "WIZ-ENROLL-001"
    mdlPersonnelEvents.SetPersonnelWizardValue "basis_text", "Wizard test enrollment"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_fio", "Wizard Test Employee"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_personal_number", "WIZ-001"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_rank", "Private"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_position", "Initial position"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_section", "Initial section"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_military_unit", "Test unit"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_vus", "100100"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_tariff_rank", "5"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_position_salary", "25000"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_rank_salary", "10000"
    enrollmentID = mdlPersonnelEvents.SavePersonnelEventInput(False)
    employeeID = CStr(mdlPersonnelEvents.GetPersonnelWizardValue("employee_id"))
    If employeeID = "" Then Err.Raise 699, , "Enrollment did not create EmployeeID"
    Set searchMatches = mdlPersonnelEvents.SearchPersonnelEmployees("Wizard Test")
    If searchMatches.Count <> 1 Then Err.Raise 699, , "Search by FIO did not return the employee"
    Set searchMatches = mdlPersonnelEvents.SearchPersonnelEmployees("WIZ-001")
    If searchMatches.Count <> 1 Then Err.Raise 699, , "Search by personal number did not return the employee"

    mdlPersonnelEvents.ResetPersonnelEventInput
    mdlPersonnelEvents.SetPersonnelWizardValue "event_type", "ENROLLMENT"
    mdlPersonnelEvents.SetPersonnelWizardValue "event_date", DateSerial(2026, 7, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "effective_date", DateSerial(2026, 7, 1)
    mdlPersonnelEvents.SetPersonnelWizardValue "order_reference", "WIZ-ENROLL-002"
    mdlPersonnelEvents.SetPersonnelWizardValue "basis_text", "Wizard duplicate search enrollment"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_fio", "Wizard Test Duplicate"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_personal_number", "WIZ-002"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_rank", "Private"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_position", "Duplicate position"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_section", "Duplicate section"
    mdlPersonnelEvents.SetPersonnelWizardValue "new_military_unit", "Test unit"
    duplicateEnrollmentID = mdlPersonnelEvents.SavePersonnelEventInput(False)
    If duplicateEnrollmentID = "" Then Err.Raise 699, , "Duplicate search fixture was not created"
    Set searchMatches = mdlPersonnelEvents.SearchPersonnelEmployees("Wizard Test")
    If searchMatches.Count <> 2 Then Err.Raise 699, , "Ambiguous search did not return exactly two employees"

    mdlPersonnelEvents.PrepareNewPersonnelAction "TRANSFER"
    Load frmPersonnelActionWizard
    If frmPersonnelActionWizard.Controls("txt_employee_id") Is Nothing Then Err.Raise 700, , "Employee field missing"
    If frmPersonnelActionWizard.Controls("txt_search") Is Nothing Then Err.Raise 700, , "Search field missing"
    frmPersonnelActionWizard.Controls("txt_search").Value = "Wizard Test"
    If InStr(1, CStr(frmPersonnelActionWizard.Controls("txt_search_results").Value), "Wizard Test", vbTextCompare) = 0 Then Err.Raise 700, , "Search preview did not show the employee"
    If frmPersonnelActionWizard.Controls("txt_destination_location") Is Nothing Then Err.Raise 701, , "Destination field missing"
    If frmPersonnelActionWizard.Controls("txt_status") Is Nothing Then Err.Raise 702, , "Status field missing"
    If frmPersonnelActionWizard.btnImportResponse.Caption = "" Then Err.Raise 703, , "Save button caption missing"

    outputPath = frmPersonnelActionWizard.ExportAction()
    If outputPath <> "" Then Err.Raise 703, , "Export before save unexpectedly returned an output path"
    If Len(Trim`$(CStr(frmPersonnelActionWizard.Controls("txt_status").Value))) = 0 Then Err.Raise 703, , "Export before save did not show a guard message"

    frmPersonnelActionWizard.Controls("txt_search").Value = "Wizard Test"
    If InStr(1, CStr(frmPersonnelActionWizard.Controls("txt_search_results").Value), "WIZ-001", vbTextCompare) = 0 Then Err.Raise 703, , "Ambiguous search preview omitted the first employee"
    If InStr(1, CStr(frmPersonnelActionWizard.Controls("txt_search_results").Value), "WIZ-002", vbTextCompare) = 0 Then Err.Raise 703, , "Ambiguous search preview omitted the second employee"
    If Trim`$(CStr(frmPersonnelActionWizard.Controls("txt_employee_id").Value)) <> "" Then Err.Raise 703, , "Ambiguous search auto-selected an employee"

    frmPersonnelActionWizard.Controls("txt_employee_id").Value = employeeID
    mdlPersonnelEvents.SetPersonnelWizardValue "employee_id", employeeID
    If Not mdlPersonnelEvents.LoadPersonnelWizardCurrentState() Then Err.Raise 704, , "Wizard could not load current state"
    frmPersonnelActionWizard.Controls("txt_event_date").Value = "02.07.2026"
    frmPersonnelActionWizard.Controls("txt_effective_date").Value = "03.07.2026"
    frmPersonnelActionWizard.Controls("txt_order_reference").Value = "WIZ-TRANSFER-001"
    frmPersonnelActionWizard.Controls("txt_basis_text").Value = "Wizard test transfer"
    frmPersonnelActionWizard.Controls("txt_new_position").Value = "Transferred position"
    frmPersonnelActionWizard.Controls("txt_new_section").Value = "Transferred section"
    frmPersonnelActionWizard.Controls("txt_new_military_unit").Value = "Test unit 2"
    frmPersonnelActionWizard.Controls("txt_new_vus").Value = "200200"
    frmPersonnelActionWizard.Controls("txt_handover_date").Value = "02.07.2026"
    frmPersonnelActionWizard.Controls("txt_acceptance_date").Value = "03.07.2026"
    frmPersonnelActionWizard.Controls("txt_duty_start_date").Value = "04.07.2026"
    transferID = frmPersonnelActionWizard.SaveAction()
    If transferID = "" Then Err.Raise 705, , "Wizard save did not return EventID"
    Set currentState = mdlPersonnelEvents.GetCurrentPersonnelState(employeeID)
    If CStr(currentState("position")) <> "Transferred position" Then Err.Raise 706, , "Wizard save did not update current state"
    outputPath = frmPersonnelActionWizard.ExportAction()
    If outputPath = "" Then Err.Raise 707, , "Wizard export did not return output path"
    transferPath = outputPath
    Unload frmPersonnelActionWizard

    mdlPersonnelEvents.PrepareNewPersonnelAction "EXCLUSION"
    Load frmPersonnelActionWizard
    If frmPersonnelActionWizard.Controls("txt_employee_id") Is Nothing Then Err.Raise 708, , "Exclusion employee field missing"
    frmPersonnelActionWizard.Controls("txt_employee_id").Value = employeeID
    mdlPersonnelEvents.SetPersonnelWizardValue "employee_id", employeeID
    If Not mdlPersonnelEvents.LoadPersonnelWizardCurrentState() Then Err.Raise 709, , "Exclusion wizard could not load current state"
    If frmPersonnelActionWizard.Controls("txt_material_assistance_status") Is Nothing Then Err.Raise 709, , "Material assistance field missing"
    If frmPersonnelActionWizard.Controls("txt_main_leave_status") Is Nothing Then Err.Raise 709, , "Main leave field missing"
    If frmPersonnelActionWizard.Controls("txt_additional_leave_status") Is Nothing Then Err.Raise 709, , "Additional leave field missing"
    If frmPersonnelActionWizard.Controls("txt_additional_leave_status").Top + frmPersonnelActionWizard.Controls("txt_additional_leave_status").Height >= frmPersonnelActionWizard.btnExportRequest.Top Then Err.Raise 709, , "Exclusion service fields overlap the action buttons"
    If frmPersonnelActionWizard.Controls("txt_status").Top + frmPersonnelActionWizard.Controls("txt_status").Height >= frmPersonnelActionWizard.btnExportRequest.Top Then Err.Raise 709, , "Exclusion status field overlaps the action buttons"
    frmPersonnelActionWizard.Controls("txt_event_date").Value = "05.07.2026"
    frmPersonnelActionWizard.Controls("txt_effective_date").Value = "06.07.2026"
    frmPersonnelActionWizard.Controls("txt_order_reference").Value = "WIZ-EXCLUSION-001"
    frmPersonnelActionWizard.Controls("txt_basis_text").Value = "Wizard test exclusion"
    frmPersonnelActionWizard.Controls("txt_handover_date").Value = "05.07.2026"
    frmPersonnelActionWizard.Controls("txt_destination_unit").Value = "Destination unit"
    frmPersonnelActionWizard.Controls("txt_destination_location").Value = "Destination city"
    frmPersonnelActionWizard.Controls("txt_material_assistance_status").Value = "оказана"
    frmPersonnelActionWizard.Controls("txt_main_leave_status").Value = "использован"
    frmPersonnelActionWizard.Controls("txt_additional_leave_status").Value = "не использован"
    exclusionID = frmPersonnelActionWizard.SaveAction()
    If exclusionID = "" Then Err.Raise 710, , "Exclusion wizard save did not return EventID"
    For employeeRow = 2 To ThisWorkbook.Worksheets("Employees").Cells(ThisWorkbook.Worksheets("Employees").Rows.Count, 1).End(xlUp).Row
        If CStr(ThisWorkbook.Worksheets("Employees").Cells(employeeRow, 1).Value) = employeeID Then Exit For
    Next employeeRow
    If CStr(ThisWorkbook.Worksheets("Employees").Cells(employeeRow, 10).Value) <> "NO" Then Err.Raise 711, , "Exclusion wizard did not deactivate employee"
    outputPath = frmPersonnelActionWizard.ExportAction()
    If outputPath = "" Then Err.Raise 712, , "Exclusion wizard export did not return output path"
    exclusionPath = outputPath
    Unload frmPersonnelActionWizard
    ProbePersonnelActionWizard = "OK|PRE_SAVE_EXPORT_BLOCKED|AMBIGUOUS_SEARCH_2|" & transferPath & "|" & exclusionPath
    Exit Function
Failed:
    ProbePersonnelActionWizard = "FAILED: " & Err.Description
End Function
"@)
    $result = $excel.Run("'$($workbook.Name)'!ProbePersonnelActionWizard")
    if ($result -notlike "OK|PRE_SAVE_EXPORT_BLOCKED|AMBIGUOUS_SEARCH_2|*") { throw $result }
    $resultParts = $result -split '\|', 5
    if ($resultParts.Count -ne 5) { throw "Unexpected personnel probe result: $result" }
    $transferPath = $resultParts[3]
    $exclusionPath = $resultParts[4]
    if (-not (Test-Path -LiteralPath $transferPath)) { throw "Transfer DOCX missing: $transferPath" }
    if (-not (Test-Path -LiteralPath $exclusionPath)) { throw "Exclusion DOCX missing: $exclusionPath" }

    $transferText = Get-DocxText $transferPath
    $exclusionText = Get-DocxText $exclusionPath
    foreach ($textValue in @($transferText, $exclusionText)) {
        if ($textValue -like '*20__*' -or $textValue -like '*№ ____*') { throw "Personnel DOCX leaked header placeholders." }
        Assert-DocxContains $textValue "§ 1" "Personnel DOCX omitted the section marker."
        Assert-DocxContains $textValue "ОСНОВАНИЕ:" "Personnel DOCX omitted the basis block."
        Assert-DocxContains $textValue "ВРИО КОМАНДИРА ВОЙСКОВОЙ ЧАСТИ" "Personnel DOCX omitted the signatory position."
        Assert-DocxContains $textValue "майор Е.Коропец" "Personnel DOCX omitted the signatory rank/name."
    }
    Assert-DocxContains $transferText "02.07.2026 г. № WIZ-TRANSFER-001" "Transfer DOCX omitted the actual order date/number."
    Assert-DocxContains $transferText "Initial position" "Transfer DOCX omitted the previous position."
    Assert-DocxContains $transferText "Transferred position" "Transfer DOCX omitted the new position."
    Assert-DocxContains $transferText "Transferred section" "Transfer DOCX omitted the new section."
    Assert-DocxContains $transferText "Test unit 2" "Transfer DOCX omitted the new military unit."
    Assert-DocxContains $transferText "ВУС-200200" "Transfer DOCX omitted the new VUS."
    Assert-DocxContains $transferText "03.07.2026" "Transfer DOCX omitted the acceptance date."
    Assert-DocxContains $transferText "04.07.2026" "Transfer DOCX omitted the duty start date."
    Assert-DocxContains $transferText "25000" "Transfer DOCX omitted the position salary."
    Assert-DocxContains $transferText "10000" "Transfer DOCX omitted the rank salary."
    Assert-DocxContains $exclusionText "05.07.2026 г. № WIZ-EXCLUSION-001" "Exclusion DOCX omitted the actual order date/number."
    Assert-DocxContains $exclusionText "Destination unit" "Exclusion DOCX omitted the destination unit."
    Assert-DocxContains $exclusionText "Destination city" "Exclusion DOCX omitted the destination location."
    Assert-DocxContains $exclusionText "Материальная помощь за текущий год: оказана" "Exclusion DOCX omitted material assistance status."
    Assert-DocxContains $exclusionText "Основной отпуск за текущий год: использован" "Exclusion DOCX omitted main leave status."
    Assert-DocxContains $exclusionText "Дополнительный отпуск за текущий год: не использован" "Exclusion DOCX omitted additional leave status."
    $workbook.Close($false); $workbook = $null
    $excel.Quit(); $excel = $null
    Write-Output "Personnel action wizard safe acceptance passed: export guard/search checks and complete transfer/exclusion DOCX structure verified."
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
