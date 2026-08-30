[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-CodeText([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    try { return ([Text.UTF8Encoding]::new($false, $true)).GetString($bytes) } catch { return [Text.Encoding]::GetEncoding(1251).GetString($bytes) }
}

function Import-CodeModuleText([object]$Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = [regex]::Replace((Read-CodeText $ModulePath), '(?m)^Attribute .*(?:\r?\n|$)', '')
    $component = $null
    try { $component = $Workbook.VBProject.VBComponents.Item($ModuleName) } catch {}
    if ($null -eq $component) { $component = $Workbook.VBProject.VBComponents.Add(1); $component.Name = $ModuleName }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    [void]$module.AddFromString($code)
}

function Import-UserForm([object]$Workbook, [string]$FormName, [string]$FormPath) {
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item($FormName)) } catch {}
    $component = $Workbook.VBProject.VBComponents.Import($FormPath)
    if ($component.Name -ne $FormName -or $component.Type -ne 3) { throw "Unexpected imported form: $($component.Name)/$($component.Type)" }
}

function Add-Probe([object]$Workbook) {
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item('grouped_order_probe')) } catch {}
    $probe = $Workbook.VBProject.VBComponents.Add(1)
    $probe.Name = 'grouped_order_probe'
    $probeCode = @'
Option Explicit

Private Function LastRow(ByVal sheetName As String) As Long
    LastRow = ThisWorkbook.Worksheets(sheetName).Cells(ThisWorkbook.Worksheets(sheetName).Rows.Count, 1).End(xlUp).Row
End Function

Private Sub ClearRows(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If LastRow(sheetName) >= 2 Then ws.Range("A2:AZ" & CStr(LastRow(sheetName))).ClearContents
End Sub

Private Sub SetCell(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnNumber As Long, ByVal valueText As String)
    ThisWorkbook.Worksheets(sheetName).Cells(rowNumber, columnNumber).Value = valueText
End Sub

Private Sub SeedFixture()
    Dim rowNumber As Long
    mdlPersonnelEvents.EnsurePersonnelEventInfrastructure
    ClearRows "Employees"
    ClearRows "EmployeeCurrentState"
    ClearRows "PersonnelEvents"
    ClearRows "PersonnelStateSnapshots"
    ClearRows "PaymentAssignments"
    ClearRows "DocumentRegistry"
    ClearRows "LegalActs"

    SetCell "Employees", 2, 1, "E1": SetCell "Employees", 2, 2, "Иванов Иван Иванович": SetCell "Employees", 2, 3, "PN-1"
    SetCell "Employees", 3, 1, "E2": SetCell "Employees", 3, 2, "Петров Петр Петрович": SetCell "Employees", 3, 3, "PN-2"
    SetCell "Employees", 4, 1, "E3": SetCell "Employees", 4, 2, "Сидоров Сидор Сидорович": SetCell "Employees", 4, 3, "PN-3"
    SetCell "Employees", 5, 1, "E4": SetCell "Employees", 5, 2, "Орлов Олег Олегович": SetCell "Employees", 5, 3, "PN-4"
    SetCell "Employees", 6, 1, "E5": SetCell "Employees", 6, 2, "Кузнецов Кузьма Кузьмич": SetCell "Employees", 6, 3, "PN-5"

    SetCell "PersonnelEvents", 2, 1, "EV1": SetCell "PersonnelEvents", 2, 2, "E1": SetCell "PersonnelEvents", 2, 3, "ENROLLMENT": SetCell "PersonnelEvents", 2, 4, "2026-08-01": SetCell "PersonnelEvents", 2, 5, "2026-08-01": SetCell "PersonnelEvents", 2, 6, "SAVED": SetCell "PersonnelEvents", 2, 8, "S1": SetCell "PersonnelEvents", 2, 9, "ORD-G-1": SetCell "PersonnelEvents", 2, 10, "Основание Иванов"
    SetCell "PersonnelEvents", 3, 1, "EV2": SetCell "PersonnelEvents", 3, 2, "E2": SetCell "PersonnelEvents", 3, 3, "TRANSFER": SetCell "PersonnelEvents", 3, 4, "2026-08-02": SetCell "PersonnelEvents", 3, 5, "2026-08-03": SetCell "PersonnelEvents", 3, 6, "SAVED": SetCell "PersonnelEvents", 3, 7, "S2": SetCell "PersonnelEvents", 3, 8, "S3": SetCell "PersonnelEvents", 3, 9, "ORD-G-2": SetCell "PersonnelEvents", 3, 10, "Основание Петров"
    SetCell "PersonnelEvents", 4, 1, "EV3": SetCell "PersonnelEvents", 4, 2, "E3": SetCell "PersonnelEvents", 4, 3, "ENROLLMENT": SetCell "PersonnelEvents", 4, 4, "2026-08-04": SetCell "PersonnelEvents", 4, 5, "2026-08-04": SetCell "PersonnelEvents", 4, 6, "SAVED": SetCell "PersonnelEvents", 4, 8, "S4": SetCell "PersonnelEvents", 4, 9, "ORD-G-3": SetCell "PersonnelEvents", 4, 10, "Основание Сидоров"
    SetCell "PersonnelEvents", 5, 1, "EV4": SetCell "PersonnelEvents", 5, 2, "E4": SetCell "PersonnelEvents", 5, 3, "EXCLUSION": SetCell "PersonnelEvents", 5, 4, "2026-08-05": SetCell "PersonnelEvents", 5, 5, "2026-08-06": SetCell "PersonnelEvents", 5, 6, "SAVED": SetCell "PersonnelEvents", 5, 7, "S5": SetCell "PersonnelEvents", 5, 9, "ORD-G-4": SetCell "PersonnelEvents", 5, 10, "Основание Орлов"
    SetCell "PersonnelEvents", 6, 1, "EV5": SetCell "PersonnelEvents", 6, 2, "E5": SetCell "PersonnelEvents", 6, 3, "TRANSFER": SetCell "PersonnelEvents", 6, 4, "2026-08-07": SetCell "PersonnelEvents", 6, 5, "2026-08-08": SetCell "PersonnelEvents", 6, 6, "SAVED": SetCell "PersonnelEvents", 6, 7, "S6": SetCell "PersonnelEvents", 6, 8, "S7": SetCell "PersonnelEvents", 6, 9, "ORD-G-5": SetCell "PersonnelEvents", 6, 10, "Основание Кузнецов"

    SetCell "PersonnelStateSnapshots", 2, 1, "S1": SetCell "PersonnelStateSnapshots", 2, 2, "EV1": SetCell "PersonnelStateSnapshots", 2, 3, "AFTER": SetCell "PersonnelStateSnapshots", 2, 4, "E1": SetCell "PersonnelStateSnapshots", 2, 6, "Должность Иванова": SetCell "PersonnelStateSnapshots", 2, 8, "Часть 1"
    SetCell "PersonnelStateSnapshots", 3, 1, "S2": SetCell "PersonnelStateSnapshots", 3, 2, "EV2": SetCell "PersonnelStateSnapshots", 3, 3, "BEFORE": SetCell "PersonnelStateSnapshots", 3, 4, "E2": SetCell "PersonnelStateSnapshots", 3, 6, "Старая должность": SetCell "PersonnelStateSnapshots", 3, 8, "Часть 1"
    SetCell "PersonnelStateSnapshots", 4, 1, "S3": SetCell "PersonnelStateSnapshots", 4, 2, "EV2": SetCell "PersonnelStateSnapshots", 4, 3, "AFTER": SetCell "PersonnelStateSnapshots", 4, 4, "E2": SetCell "PersonnelStateSnapshots", 4, 6, "Новая должность": SetCell "PersonnelStateSnapshots", 4, 8, "Часть 2"
    SetCell "PersonnelStateSnapshots", 5, 1, "S4": SetCell "PersonnelStateSnapshots", 5, 2, "EV3": SetCell "PersonnelStateSnapshots", 5, 3, "AFTER": SetCell "PersonnelStateSnapshots", 5, 4, "E3": SetCell "PersonnelStateSnapshots", 5, 6, "Должность Сидорова": SetCell "PersonnelStateSnapshots", 5, 8, "Часть 1"
    SetCell "PersonnelStateSnapshots", 6, 1, "S5": SetCell "PersonnelStateSnapshots", 6, 2, "EV4": SetCell "PersonnelStateSnapshots", 6, 3, "BEFORE": SetCell "PersonnelStateSnapshots", 6, 4, "E4": SetCell "PersonnelStateSnapshots", 6, 6, "Должность Орлова": SetCell "PersonnelStateSnapshots", 6, 8, "Часть 1"
    SetCell "PersonnelStateSnapshots", 7, 1, "S6": SetCell "PersonnelStateSnapshots", 7, 2, "EV5": SetCell "PersonnelStateSnapshots", 7, 3, "BEFORE": SetCell "PersonnelStateSnapshots", 7, 4, "E5": SetCell "PersonnelStateSnapshots", 7, 6, "Старая должность Кузнецова": SetCell "PersonnelStateSnapshots", 7, 8, "Часть 1"
    SetCell "PersonnelStateSnapshots", 8, 1, "S7": SetCell "PersonnelStateSnapshots", 8, 2, "EV5": SetCell "PersonnelStateSnapshots", 8, 3, "AFTER": SetCell "PersonnelStateSnapshots", 8, 4, "E5": SetCell "PersonnelStateSnapshots", 8, 6, "Новая должность Кузнецова": SetCell "PersonnelStateSnapshots", 8, 8, "Часть 2"

    SetCell "PaymentAssignments", 2, 1, "A1": SetCell "PaymentAssignments", 2, 2, "E1": SetCell "PaymentAssignments", 2, 3, "EV1": SetCell "PaymentAssignments", 2, 5, "FIZO": SetCell "PaymentAssignments", 2, 6, "PERCENT": SetCell "PaymentAssignments", 2, 7, "70": SetCell "PaymentAssignments", 2, 11, "ACTIVE": SetCell "PaymentAssignments", 2, 13, "LA1": SetCell "PaymentAssignments", 2, 15, "Результаты ФИЗО"
    SetCell "PaymentAssignments", 3, 1, "A2": SetCell "PaymentAssignments", 3, 2, "E1": SetCell "PaymentAssignments", 3, 3, "EV1": SetCell "PaymentAssignments", 3, 5, "DRIVER_C_D_CE": SetCell "PaymentAssignments", 3, 6, "PERCENT": SetCell "PaymentAssignments", 3, 7, "50": SetCell "PaymentAssignments", 3, 11, "ACTIVE": SetCell "PaymentAssignments", 3, 13, "LA1": SetCell "PaymentAssignments", 3, 15, "Водительское удостоверение"
    SetCell "PaymentAssignments", 4, 1, "A3": SetCell "PaymentAssignments", 4, 2, "E2": SetCell "PaymentAssignments", 4, 3, "EV2": SetCell "PaymentAssignments", 4, 5, "FIZO": SetCell "PaymentAssignments", 4, 6, "PERCENT": SetCell "PaymentAssignments", 4, 7, "50": SetCell "PaymentAssignments", 4, 11, "ACTIVE": SetCell "PaymentAssignments", 4, 13, "LA1": SetCell "PaymentAssignments", 4, 15, "Результаты ФИЗО Петрова"
    SetCell "PaymentAssignments", 5, 1, "A4": SetCell "PaymentAssignments", 5, 2, "E4": SetCell "PaymentAssignments", 5, 3, "EV4": SetCell "PaymentAssignments", 5, 5, "MEDAL": SetCell "PaymentAssignments", 5, 6, "PERCENT": SetCell "PaymentAssignments", 5, 7, "20": SetCell "PaymentAssignments", 5, 11, "ACTIVE": SetCell "PaymentAssignments", 5, 13, "LA1": SetCell "PaymentAssignments", 5, 15, "Наградной лист"
    SetCell "PaymentAssignments", 6, 1, "A5": SetCell "PaymentAssignments", 6, 2, "E5": SetCell "PaymentAssignments", 6, 3, "EV5": SetCell "PaymentAssignments", 6, 5, "TARIFF_1_4": SetCell "PaymentAssignments", 6, 6, "PERCENT": SetCell "PaymentAssignments", 6, 7, "40": SetCell "PaymentAssignments", 6, 11, "ACTIVE": SetCell "PaymentAssignments", 6, 13, "LA1": SetCell "PaymentAssignments", 6, 15, "Штатное расписание"
    SetCell "LegalActs", 2, 1, "LA1": SetCell "LegalActs", 2, 2, "ORDER": SetCell "LegalActs", 2, 3, "1": SetCell "LegalActs", 2, 4, "2026-01-01": SetCell "LegalActs", 2, 5, "Fixture legal act"
End Sub

Private Function CountText(ByVal sourceText As String, ByVal needle As String) As Long
    Dim at As Long
    Dim startAt As Long
    startAt = 1
    Do
        at = InStr(startAt, sourceText, needle, vbTextCompare)
        If at = 0 Then Exit Do
        CountText = CountText + 1
        startAt = at + Len(needle)
    Loop
End Function

Private Function RowCount(ByVal sheetName As String) As Long
    RowCount = LastRow(sheetName) - 1
    If RowCount < 0 Then RowCount = 0
End Function

Public Function RunGroupedProbe() As String
    Dim report As String
    Dim validNoPayment As String
    Dim outputPath As String
    Dim invalidExportError As Long
    Dim docsBefore As Long
    Dim linksBefore As Long
    Dim docsAfterInvalid As Long
    Dim linksAfterInvalid As Long
    Dim decisionReport As String
    Dim ruleRow As Long
    On Error GoTo Failed
    SeedFixture
    report = mdlGroupedPersonnelOrderExport.BuildGroupedPersonnelOrderReport()
    If Left$(report, 3) <> "OK|" Then Err.Raise vbObjectError + 880, , "Expected OK report: " & report
    If InStr(1, report, "paragraph_count=3", vbTextCompare) = 0 Then Err.Raise vbObjectError + 881, , "Expected three category paragraphs."
    If CountText(report, "ITEM|") <> 5 Then Err.Raise vbObjectError + 882, , "Expected five employee items."
    If CountText(report, "PAYMENT|") <> 5 Then Err.Raise vbObjectError + 883, , "Expected five independent payments."
    If CountText(report, "payment_code=FIZO") <> 2 Then Err.Raise vbObjectError + 884, , "FIZO payments were merged or lost."
    If InStr(1, report, "PARAGRAPH|no=1|event_type=ENROLLMENT", vbTextCompare) = 0 Then Err.Raise vbObjectError + 885, , "Enrollment paragraph is missing."
    If InStr(1, report, "PARAGRAPH|no=2|event_type=TRANSFER", vbTextCompare) = 0 Then Err.Raise vbObjectError + 886, , "Transfer paragraph is missing."
    If InStr(1, report, "PARAGRAPH|no=3|event_type=EXCLUSION", vbTextCompare) = 0 Then Err.Raise vbObjectError + 887, , "Exclusion paragraph is missing."
    If InStr(1, report, "WARNING|event_id=EV1", vbTextCompare) = 0 Then Err.Raise vbObjectError + 888, , "Driver-over-cap warning is missing."

    validNoPayment = mdlGroupedPersonnelOrderExport.ValidateGroupedPersonnelOrder("EV3")
    If Left$(validNoPayment, 6) <> "VALID|" Then Err.Raise vbObjectError + 889, , "Unchecked payment incorrectly blocked export: " & validNoPayment

    ruleRow = LastRow("PaymentRules") + 1
    SetCell "PaymentRules", ruleRow, 1, "RULE-REQUIRES-DECISION"
    SetCell "PaymentRules", ruleRow, 2, "TEST_REQUIRES_DECISION"
    SetCell "PaymentRules", ruleRow, 22, "REQUIRES_DECISION"
    ThisWorkbook.Worksheets("PaymentAssignments").Cells(4, 5).Value = "TEST_REQUIRES_DECISION"
    decisionReport = mdlGroupedPersonnelOrderExport.ValidateGroupedPersonnelOrder("EV2")
    If Left$(decisionReport, 8) <> "INVALID|" Then Err.Raise vbObjectError + 899, , "REQUIRES_DECISION rule unexpectedly passed validation."
    If InStr(1, decisionReport, "requires an explicit legal decision", vbTextCompare) = 0 Then Err.Raise vbObjectError + 900, , "REQUIRES_DECISION reason was not reported."
    ThisWorkbook.Worksheets("PaymentAssignments").Cells(4, 5).Value = "FIZO"

    docsBefore = RowCount("DocumentRegistry")
    linksBefore = 0
    On Error Resume Next
    linksBefore = RowCount("DocumentEventLinks")
    On Error GoTo Failed
    ThisWorkbook.Worksheets("PaymentAssignments").Cells(2, 15).Value = vbNullString
    If Left$(mdlGroupedPersonnelOrderExport.ValidateGroupedPersonnelOrder("EV1"), 8) <> "INVALID|" Then Err.Raise vbObjectError + 890, , "Missing factual basis did not block validation."
    If InStr(1, mdlGroupedPersonnelOrderExport.ValidateGroupedPersonnelOrder("EV1"), "field=factual_basis", vbTextCompare) = 0 Then Err.Raise vbObjectError + 891, , "Missing factual-basis field was not reported."
    On Error Resume Next
    outputPath = mdlGroupedPersonnelOrderExport.ExportGroupedPersonnelOrder("EV1")
    invalidExportError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If invalidExportError = 0 Then Err.Raise vbObjectError + 892, , "Invalid payment unexpectedly exported."
    docsAfterInvalid = RowCount("DocumentRegistry")
    linksAfterInvalid = 0
    On Error Resume Next
    linksAfterInvalid = RowCount("DocumentEventLinks")
    On Error GoTo Failed
    If docsAfterInvalid <> docsBefore Then Err.Raise vbObjectError + 893, , "Invalid export changed DocumentRegistry."
    If linksAfterInvalid <> linksBefore Then Err.Raise vbObjectError + 894, , "Invalid export changed DocumentEventLinks."
    ThisWorkbook.Worksheets("PaymentAssignments").Cells(2, 15).Value = "Результаты ФИЗО"

    outputPath = mdlGroupedPersonnelOrderExport.ExportGroupedPersonnelOrder("EV1,EV2,EV3,EV4,EV5")
    If Len(outputPath) = 0 Then Err.Raise vbObjectError + 895, , "Valid export returned no path."
    If RowCount("DocumentRegistry") <> docsBefore + 1 Then Err.Raise vbObjectError + 896, , "Grouped document was not registered once."
    If RowCount("DocumentEventLinks") <> 5 Then Err.Raise vbObjectError + 897, , "Expected one link per event."
    If ThisWorkbook.Worksheets("PersonnelEvents").Cells(2, 6).Value <> "EXPORTED" Then Err.Raise vbObjectError + 898, , "EV1 was not marked exported."
    RunGroupedProbe = "GROUPED_OK|path=" & outputPath & "|report=" & Replace$(Replace$(Split(report, vbCrLf)(0), vbCr, ""), vbLf, "")
    Exit Function
Failed:
    RunGroupedProbe = "FAILED|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunGroupedFormProbe() As String
    On Error GoTo Failed
    Load frmGroupedPersonnelOrderWizard
    If Not frmGroupedPersonnelOrderWizard.Controls.Item("txtReport").Locked Then Err.Raise vbObjectError + 910, , "Grouped report control must be read-only."
    If Not frmGroupedPersonnelOrderWizard.Controls.Item("txtReport").MultiLine Then Err.Raise vbObjectError + 911, , "Grouped report control must be multiline."
    Unload frmGroupedPersonnelOrderWizard
    RunGroupedFormProbe = "GROUPED_FORM_OK"
    Exit Function
Failed:
    On Error Resume Next
    Unload frmGroupedPersonnelOrderWizard
    RunGroupedFormProbe = "FAILED|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@
    [void]$probe.CodeModule.AddFromString($probeCode)
}

function Get-RegistryHash([object]$Worksheet) {
    $lastRow = [int]$Worksheet.Cells.Item($Worksheet.Rows.Count, 1).End(-4162).Row
    $lastColumn = [int]$Worksheet.Cells.Item(1, $Worksheet.Columns.Count).End(-4159).Column
    $parts = [Collections.Generic.List[string]]::new()
    for ($row = 1; $row -le $lastRow; $row++) {
        for ($column = 1; $column -le $lastColumn; $column++) { [void]$parts.Add([string]$Worksheet.Cells.Item($row, $column).Value2) }
        [void]$parts.Add('<ROW>')
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`0")))).Replace('-', '')) } finally { $sha.Dispose() }
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {} }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word is running. Close Office before grouped-order test.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testDirectory = Join-Path $projectRoot "Trash\grouped-personnel-order-$stamp"
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$testWorkbook = Join-Path $testDirectory 'CreateOrder.grouped.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $testWorkbook -Force

$excel = $null
$workbook = $null
$beforeHashes = @{}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($testWorkbook, 0, $false)
    Import-CodeModuleText $workbook 'mdlGroupedPersonnelOrderExport' (Join-Path $resolvedSource 'mdlGroupedPersonnelOrderExport.bas')
    Import-CodeModuleText $workbook 'ModuleLocalization' (Join-Path $resolvedSource 'ModuleLocalization.bas')
    Import-UserForm $workbook 'frmGroupedPersonnelOrderWizard' (Join-Path $resolvedSource 'frmGroupedPersonnelOrderWizard.frm')
    Add-Probe $workbook
    foreach ($sheetName in @('Employees', 'EmployeeCurrentState', 'PersonnelEvents', 'PersonnelStateSnapshots', 'PaymentAssignments', 'DocumentRegistry', 'LegalActs')) {
        $beforeHashes[$sheetName] = Get-RegistryHash $workbook.Worksheets.Item($sheetName)
    }
    $probe = [string]$excel.Run("'$($workbook.Name)'!grouped_order_probe.RunGroupedProbe")
    if ($probe -notlike 'GROUPED_OK|*') { throw "Grouped-order probe failed: $probe" }
    $formProbe = [string]$excel.Run("'$($workbook.Name)'!grouped_order_probe.RunGroupedFormProbe")
    if ($formProbe -ne 'GROUPED_FORM_OK') { throw "Grouped-order form probe failed: $formProbe" }
    if ($probe -notmatch '\|path=(.+?)\|report=') { throw "Grouped-order probe did not return an output path: $probe" }
    $outputPath = $Matches[1]
    if (-not (Test-Path -LiteralPath $outputPath)) { throw "Grouped DOCX was not created: $outputPath" }
    Write-Output "GROUPED_PERSONNEL_ORDER_OK|$probe|form=$formProbe|workbook=$testWorkbook"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject $workbook }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
for ($attempt = 1; $attempt -le 120 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after grouped-order test.' }

$docxPath = $outputPath
$zip = $null
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($docxPath)
    $entry = $zip.GetEntry('word/document.xml')
    if ($null -eq $entry) { throw 'DOCX is missing word/document.xml.' }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    foreach ($needle in @('§1', '§2', '§3', 'Иванов', 'Петров', 'Сидоров', 'Орлов', 'Кузнецов')) { if ($xml -notmatch [regex]::Escape($needle)) { throw "DOCX is missing expected text: $needle" } }
    if (($xml -split 'FIZO').Count - 1 -ne 2) { throw 'DOCX did not preserve two separate FIZO payment lines.' }
}
finally { if ($null -ne $zip) { $zip.Dispose() } }

Write-Output "GROUPED_PERSONNEL_ORDER_DOCX_OK|$docxPath"
