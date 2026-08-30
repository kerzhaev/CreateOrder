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
    $code = Read-CodeText $ModulePath
    $code = [regex]::Replace($code, '(?m)^Attribute .*\r?\n', '')
    $component = $null
    try { $component = $Workbook.VBProject.VBComponents.Item($ModuleName) } catch {}
    if ($null -eq $component) { $component = $Workbook.VBProject.VBComponents.Add(1); $component.Name = $ModuleName }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    [void]$module.AddFromString($code)
}

function Add-Probe([object]$Workbook) {
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item('history_center_probe')) } catch {}
    $probe = $Workbook.VBProject.VBComponents.Add(1)
    $probe.Name = 'history_center_probe'
    $probeCode = @'
Option Explicit

Public Function RunHistoryProbe() As String
    Dim report As String
    Dim pathText As String
    Dim missingError As Long
    Dim ambiguousError As Long
    On Error GoTo Failed
    report = mdlPersonnelHistoryCenter.BuildPersonnelHistoryCenterReport("E1")
    pathText = mdlPersonnelHistoryCenter.GetPersonnelHistoryDocumentPath("E1", "EV1", "D1")
    On Error Resume Next
    pathText = mdlPersonnelHistoryCenter.GetPersonnelHistoryDocumentPath("E1", "EV1", "D-MISSING")
    missingError = Err.Number
    Err.Clear
    pathText = mdlPersonnelHistoryCenter.ResolvePersonnelHistoryEmployeeID("Duplicate Fixture")
    ambiguousError = Err.Number
    On Error GoTo Failed
    If InStr(1, report, "EVENT | id=EV2", vbTextCompare) = 0 Then Err.Raise vbObjectError + 780, "history_center_probe", "EV2 is missing."
    If InStr(1, report, "EVENT | id=EV1", vbTextCompare) = 0 Then Err.Raise vbObjectError + 781, "history_center_probe", "EV1 is missing."
    If InStr(1, report, "SUMMARY | events=2", vbTextCompare) = 0 Then Err.Raise vbObjectError + 782, "history_center_probe", "Summary count is wrong."
    If InStr(1, report, "DOCUMENT | id=D1", vbTextCompare) = 0 Then Err.Raise vbObjectError + 783, "history_center_probe", "Document is missing."
    If InStr(1, report, "SNAPSHOT | id=S1", vbTextCompare) = 0 Then Err.Raise vbObjectError + 784, "history_center_probe", "Snapshot is missing."
    If InStr(1, report, "ASSIGNMENT | id=A1", vbTextCompare) = 0 Then Err.Raise vbObjectError + 785, "history_center_probe", "Assignment is missing."
    If InStr(InStr(1, report, "-- EVENTS --", vbTextCompare), report, "EVENT | id=EV2", vbTextCompare) > InStr(InStr(1, report, "-- EVENTS --", vbTextCompare), report, "EVENT | id=EV1", vbTextCompare) Then Err.Raise vbObjectError + 786, "history_center_probe", "Events are not chronologically ordered."
    RunHistoryProbe = "OK|missing_error=" & CStr(missingError) & "|ambiguous_error=" & CStr(ambiguousError)
    Exit Function
Failed:
    RunHistoryProbe = "FAILED|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@
    [void]$probe.CodeModule.AddFromString($probeCode)
}

function Find-LastRow([object]$Worksheet) {
    $found = $Worksheet.Cells.Find('*', $Worksheet.Cells.Item(1, 1), -4123, 1, 1, 2, $false)
    if ($null -eq $found) { return 1 }
    return [int]$found.Row
}

function Headers([object]$Worksheet) {
    $map = @{}
    $lastColumn = $Worksheet.Cells.Item(1, $Worksheet.Columns.Count).End(-4159).Column
    for ($column = 1; $column -le $lastColumn; $column++) {
        $header = [string]$Worksheet.Cells.Item(1, $column).Value2
        if (-not [string]::IsNullOrWhiteSpace($header)) { $map[$header] = $column }
    }
    return $map
}

function Clear-DataRows([object]$Worksheet) {
    $lastRow = Find-LastRow $Worksheet
    if ($lastRow -ge 2) { [void]$Worksheet.Range("A2:AZ$lastRow").ClearContents() }
}

function Set-Row([object]$Worksheet, [int]$Row, [hashtable]$Values) {
    $map = Headers $Worksheet
    foreach ($item in $Values.GetEnumerator()) {
        if (-not $map.ContainsKey($item.Key)) { throw "Missing header $($item.Key) on $($Worksheet.Name)." }
        $null = $Worksheet.Cells.Item($Row, $map[$item.Key]).Value2 = [string]$item.Value
    }
}

function Get-SheetHash([object]$Worksheet) {
    $lastRow = Find-LastRow $Worksheet
    $lastColumn = $Worksheet.Cells.Item(1, $Worksheet.Columns.Count).End(-4159).Column
    $parts = [Collections.Generic.List[string]]::new()
    for ($row = 1; $row -le $lastRow; $row++) {
        for ($column = 1; $column -le $lastColumn; $column++) { [void]$parts.Add([string]$Worksheet.Cells.Item($row, $column).Value2) }
        [void]$parts.Add('<ROW>')
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`0")))).Replace('-', '')) } finally { $sha.Dispose() }
}

function Get-RegistryHashes([object]$Workbook) {
    $hashes = @{}
    foreach ($name in @('Employees', 'EmployeeCurrentState', 'PersonnelEvents', 'PersonnelStateSnapshots', 'PaymentAssignments', 'DocumentRegistry', 'StaffStateSyncLog', 'Штат')) {
        try { $hashes[$name] = Get-SheetHash $Workbook.Worksheets.Item($name) } catch {}
    }
    return $hashes
}

function Seed-Fixture([object]$Workbook, [string]$WorkbookFilePath) {
    foreach ($name in @('Employees', 'EmployeeCurrentState', 'PersonnelEvents', 'PersonnelStateSnapshots', 'PaymentAssignments', 'DocumentRegistry')) { Clear-DataRows $Workbook.Worksheets.Item($name) }
    Set-Row $Workbook.Worksheets.Item('Employees') 2 @{ EmployeeID='E1'; FIO='Fixture Employee'; PersonalNumber='PN-1'; TableNumber='T-1'; SourceMode='FIXTURE'; StaffLinkStatus='MANUAL_ONLY'; IsActive='YES'; CreatedAt='2026-08-01'; UpdatedAt='2026-08-01' }
    Set-Row $Workbook.Worksheets.Item('Employees') 3 @{ EmployeeID='E2'; FIO='Duplicate Fixture'; PersonalNumber='PN-2'; TableNumber='T-2'; SourceMode='FIXTURE'; StaffLinkStatus='MANUAL_ONLY'; IsActive='YES'; CreatedAt='2026-08-01'; UpdatedAt='2026-08-01' }
    Set-Row $Workbook.Worksheets.Item('Employees') 4 @{ EmployeeID='E3'; FIO='Duplicate Fixture'; PersonalNumber='PN-3'; TableNumber='T-3'; SourceMode='FIXTURE'; StaffLinkStatus='MANUAL_ONLY'; IsActive='YES'; CreatedAt='2026-08-01'; UpdatedAt='2026-08-01' }
    Set-Row $Workbook.Worksheets.Item('PersonnelEvents') 2 @{ EventID='EV1'; EmployeeID='E1'; EventType='TRANSFER'; EventDate='2026-08-03'; EffectiveDate='2026-08-03'; Status='SAVED'; BeforeSnapshotID='S1'; AfterSnapshotID='S2'; OrderReference='ORD-1'; BasisText='Fixture'; CreatedAt='2026-08-03'; UpdatedAt='2026-08-03' }
    Set-Row $Workbook.Worksheets.Item('PersonnelEvents') 3 @{ EventID='EV2'; EmployeeID='E1'; EventType='ENROLLMENT'; EventDate='2026-08-01'; EffectiveDate='2026-08-01'; Status='EXPORTED'; BeforeSnapshotID=''; AfterSnapshotID='S3'; OrderReference='ORD-2'; BasisText='Fixture'; CreatedAt='2026-08-01'; UpdatedAt='2026-08-01' }
    Set-Row $Workbook.Worksheets.Item('PersonnelStateSnapshots') 2 @{ SnapshotID='S1'; EventID='EV1'; SnapshotKind='BEFORE'; EmployeeID='E1'; Position='Old position'; StateDate='2026-08-03'; CreatedAt='2026-08-03' }
    Set-Row $Workbook.Worksheets.Item('PersonnelStateSnapshots') 3 @{ SnapshotID='S2'; EventID='EV1'; SnapshotKind='AFTER'; EmployeeID='E1'; Position='New position'; StateDate='2026-08-03'; CreatedAt='2026-08-03' }
    Set-Row $Workbook.Worksheets.Item('PersonnelStateSnapshots') 4 @{ SnapshotID='S3'; EventID='EV2'; SnapshotKind='AFTER'; EmployeeID='E1'; Position='Initial position'; StateDate='2026-08-01'; CreatedAt='2026-08-01' }
    Set-Row $Workbook.Worksheets.Item('EmployeeCurrentState') 2 @{ EmployeeID='E1'; Rank='Private'; Position='New position'; Section='Section'; MilitaryUnit='Unit'; StateDate='2026-08-03'; LastEventID='EV1' }
    Set-Row $Workbook.Worksheets.Item('PaymentAssignments') 2 @{ AssignmentID='A1'; EmployeeID='E1'; EventID='EV1'; PaymentType='Fixture'; PaymentCode='FIX'; AmountKind='PERCENT'; AmountValue='10'; StartDate='2026-08-03'; EndDate=''; Status='ACTIVE' }
    Set-Row $Workbook.Worksheets.Item('DocumentRegistry') 2 @{ DocumentID='D1'; EventID='EV1'; DocumentType='TRANSFER_ORDER'; DocumentNumber='1'; DocumentDate='2026-08-03'; FilePath=$WorkbookFilePath; Status='EXPORTED' }
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {} }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word is running. Close Office before the history center test.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testDirectory = Join-Path $projectRoot "Trash\personnel-history-center-$stamp"
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$testWorkbook = Join-Path $testDirectory 'CreateOrder.history-center.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $testWorkbook -Force

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($testWorkbook, 0, $false)
    Import-CodeModuleText $workbook 'mdlPersonnelHistoryCenter' (Join-Path $resolvedSource 'mdlPersonnelHistoryCenter.bas')
    Add-Probe $workbook
    Seed-Fixture $workbook $testWorkbook
    $before = Get-RegistryHashes $workbook
    $probe = [string]$excel.Run("'$($workbook.Name)'!history_center_probe.RunHistoryProbe")
    $after = Get-RegistryHashes $workbook
    foreach ($key in $before.Keys) { if (-not $after.ContainsKey($key) -or $before[$key] -ne $after[$key]) { throw "History center mutated sheet $key." } }
    if ($probe -notlike 'OK|*') { throw "History center probe failed: $probe" }
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null
    $excel.Quit()
    Release-ComObject $excel
    $excel = $null
    Write-Output "PERSONNEL_HISTORY_CENTER_OK|$probe|workbook=$testWorkbook"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject $workbook }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
for ($attempt = 1; $attempt -le 120 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after the history center test.' }
