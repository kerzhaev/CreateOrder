[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-VbaText {
    param([Parameter(Mandatory = $true)][string]$Path)
    [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $true)][string]$ModulePath
    )
    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $null
    try { $component = $Workbook.VBProject.VBComponents.Item($ModuleName) } catch {}
    if ($null -eq $component) {
        $component = $Workbook.VBProject.VBComponents.Add(1)
        $component.Name = $ModuleName
    }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $null = $module.AddFromString($code)
}

function Add-Probe {
    param([Parameter(Mandatory = $true)][object]$Workbook)
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item('personnel_integrity_probe')) } catch {}
    $probe = $Workbook.VBProject.VBComponents.Add(1)
    $probe.Name = 'personnel_integrity_probe'
    $probeCode = @'
Option Explicit

Public Function RunIntegrityProbe() As String
    Dim findings As Collection
    Dim item As Object
    Dim result As String
    On Error GoTo Failed
    Set findings = mdlPersonnelDataIntegrity.ScanPersonnelDataIntegrity()
    For Each item In findings
        result = result & CStr(item("Category")) & "=" & CStr(item("Severity")) & ";"
    Next item
    RunIntegrityProbe = "OK|" & CStr(findings.Count) & "|" & result & "|" & mdlPersonnelDataIntegrity.BuildPersonnelDataIntegrityReport()
    Exit Function
Failed:
    RunIntegrityProbe = "FAILED|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@
    $null = $probe.CodeModule.AddFromString($probeCode)
}

function Find-LastRow {
    param([Parameter(Mandatory = $true)][object]$Worksheet)
    $found = $Worksheet.Cells.Find('*', $Worksheet.Cells.Item(1, 1), -4123, 1, 1, 2, $false)
    if ($null -eq $found) { return 1 }
    [int]$found.Row
}

function Get-HeaderMap {
    param([Parameter(Mandatory = $true)][object]$Worksheet)
    $map = @{}
    $lastColumn = $Worksheet.Cells.Item(1, $Worksheet.Columns.Count).End(-4159).Column
    for ($column = 1; $column -le $lastColumn; $column++) {
        $name = [string]$Worksheet.Cells.Item(1, $column).Value2
        if (-not [string]::IsNullOrWhiteSpace($name)) { $map[$name] = $column }
    }
    $map
}

function Clear-DataRows {
    param([Parameter(Mandatory = $true)][object]$Worksheet)
    $lastRow = Find-LastRow -Worksheet $Worksheet
    if ($lastRow -ge 2) { $null = $Worksheet.Range("A2:AZ$lastRow").ClearContents() }
}

function Set-DataRow {
    param(
        [Parameter(Mandatory = $true)][object]$Worksheet,
        [Parameter(Mandatory = $true)][int]$Row,
        [Parameter(Mandatory = $true)][hashtable]$Values
    )
    $headers = Get-HeaderMap -Worksheet $Worksheet
    foreach ($entry in $Values.GetEnumerator()) {
        if (-not $headers.ContainsKey($entry.Key)) { throw "Missing header '$($entry.Key)' on $($Worksheet.Name)." }
        $value = $entry.Value
        if ($value -is [datetime]) { $value = $value.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) }
        $null = $Worksheet.Cells.Item($Row, $headers[$entry.Key]).Value2 = [string]$value
    }
}

function Seed-ValidFixture {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$WorkbookFilePath
    )
    $registrySheets = @('Employees', 'EmployeeCurrentState', 'PersonnelEvents', 'PersonnelStateSnapshots', 'PaymentAssignments', 'DocumentRegistry', 'LegalActs', 'PaymentRules', 'PositionClassification')
    foreach ($name in $registrySheets) { Clear-DataRows -Worksheet $Workbook.Worksheets.Item($name) }

    $eventDate = [datetime]::Parse('2026-08-01T00:00:00')
    $created = [datetime]::Parse('2026-08-01T12:00:00')
    $snapshotCreated = [datetime]::Parse('2026-08-01T12:01:00')
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('Employees') -Row 2 -Values @{
        EmployeeID='E1'; FIO='Fixture Employee'; PersonalNumber='PN-1'; TableNumber='T-1'; SourceMode='FIXTURE'; StaffLinkStatus='MANUAL_ONLY'; StaffReference=''; CreatedAt=$created; UpdatedAt=$created; IsActive='YES'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PersonnelEvents') -Row 2 -Values @{
        EventID='EV1'; EmployeeID='E1'; EventType='ENROLLMENT'; EventDate=$eventDate; EffectiveDate=$eventDate; Status='DRAFT'; BeforeSnapshotID='S1'; AfterSnapshotID='S2'; OrderReference='FIX-1'; BasisText='Fixture'; OperatorName='test'; CreatedAt=$created; UpdatedAt=$created
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PersonnelStateSnapshots') -Row 2 -Values @{
        SnapshotID='S1'; EventID='EV1'; SnapshotKind='BEFORE'; EmployeeID='E1'; Rank='Private'; Position='Position'; Section='Section'; MilitaryUnit='Unit'; StateDate=$eventDate; CreatedAt=$snapshotCreated
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PersonnelStateSnapshots') -Row 3 -Values @{
        SnapshotID='S2'; EventID='EV1'; SnapshotKind='AFTER'; EmployeeID='E1'; Rank='Private'; Position='Position'; Section='Section'; MilitaryUnit='Unit'; StateDate=$eventDate; CreatedAt=$snapshotCreated
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('EmployeeCurrentState') -Row 2 -Values @{
        EmployeeID='E1'; Rank='Private'; Position='Position'; Section='Section'; MilitaryUnit='Unit'; StateDate=$eventDate; SourceEventID='EV1'; LastEventID='EV1'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PaymentAssignments') -Row 2 -Values @{
        AssignmentID='A1'; EmployeeID='E1'; EventID='EV1'; PaymentType='Fixture'; PaymentCode='FIX'; AmountKind='PERCENT'; AmountValue=10; StartDate=$eventDate; Status='ACTIVE'; ActID='ACT1'; ActPoint='1'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('DocumentRegistry') -Row 2 -Values @{
        DocumentID='D1'; EventID='EV1'; DocumentType='FIXTURE'; DocumentNumber='1'; DocumentDate=$eventDate; FilePath=$WorkbookFilePath; Status='EXPORTED'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('LegalActs') -Row 2 -Values @{
        ActID='ACT1'; ActType='ORDER'; ActNumber='1'; ActDate=$eventDate; Title='Fixture'; EffectiveFrom=$eventDate; AccessMark='CONFIRMED'
    }
}

function Add-CorruptFixture {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$WorkbookFilePath
    )
    $employees = $Workbook.Worksheets.Item('Employees')
    Set-DataRow -Worksheet $employees -Row 3 -Values @{
        EmployeeID='E1'; FIO='Fixture Duplicate'; PersonalNumber='PN-2'; TableNumber='T-2'; SourceMode='FIXTURE'; StaffLinkStatus='MANUAL_ONLY'; StaffReference=''; CreatedAt='2026-08-01'; UpdatedAt='2026-08-01'; IsActive='YES'
    }
    Set-DataRow -Worksheet $employees -Row 4 -Values @{
        EmployeeID='E-LINK'; FIO='Fixture Linked'; PersonalNumber='PN-LINK'; SourceMode='FIXTURE'; StaffLinkStatus='LINKED'; StaffReference='STAFF_ROW:999'; CreatedAt='2026-08-01'; UpdatedAt='2026-08-01'; IsActive='YES'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PersonnelEvents') -Row 3 -Values @{
        EventID='EV-BAD-EMP'; EmployeeID='E-MISSING'; EventType='TRANSFER'; EventDate='2026-08-02'; EffectiveDate='2026-08-02'; Status='DRAFT'; CreatedAt='2026-08-02'; UpdatedAt='2026-08-02'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PersonnelEvents') -Row 4 -Values @{
        EventID='EV-MISSING-SNAPSHOT'; EmployeeID='E1'; EventType='TRANSFER'; EventDate='2026-08-03'; EffectiveDate='2026-08-03'; Status='DRAFT'; CreatedAt='2026-08-03'; UpdatedAt='2026-08-03'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PersonnelEvents') -Row 5 -Values @{
        EventID='EV-EXCLUSION'; EmployeeID='E1'; EventType='EXCLUSION'; EventDate='2026-08-04'; EffectiveDate='2026-08-04'; Status='VERIFIED'; CreatedAt='2026-08-04'; UpdatedAt='2026-08-04'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PersonnelStateSnapshots') -Row 4 -Values @{
        SnapshotID='S-BAD-KIND'; EventID='EV1'; SnapshotKind='BROKEN'; EmployeeID='E1'; StateDate='2026-08-01'; CreatedAt='2026-08-01T12:02:00'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('EmployeeCurrentState') -Row 3 -Values @{
        EmployeeID='E-MISSING-STATE'; StateDate='2026-08-02'; SourceEventID='EV-MISSING-STATE'; LastEventID='EV-MISSING-STATE'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('PaymentAssignments') -Row 3 -Values @{
        AssignmentID='A-BAD'; EmployeeID='E-MISSING'; EventID='EV-MISSING'; PaymentType='Fixture'; PaymentCode='BAD'; AmountKind='PERCENT'; AmountValue=10; StartDate='2026-08-02'; Status='ACTIVE'; ActID='ACT-MISSING'
    }
    Set-DataRow -Worksheet $Workbook.Worksheets.Item('DocumentRegistry') -Row 3 -Values @{
        DocumentID='D-BAD'; EventID='EV-MISSING-DOC'; DocumentType='FIXTURE'; FilePath='Z:\missing\fixture.docx'; Status='EXPORTED'
    }
    $events = $Workbook.Worksheets.Item('PersonnelEvents')
    $eventHeaders = Get-HeaderMap -Worksheet $events
    $null = $events.Cells.Item(2, $eventHeaders['EffectiveDate']).Value2 = '2026-07-31 00:00:00'
    $null = $Workbook.Worksheets.Item('Employees').Cells.Item(2, (Get-HeaderMap -Worksheet $employees)['IsActive']).Value2 = 'NO'
}

function Get-SheetHash {
    param([Parameter(Mandatory = $true)][object]$Worksheet)
    $lastRow = Find-LastRow -Worksheet $Worksheet
    $lastColumn = $Worksheet.Cells.Item(1, $Worksheet.Columns.Count).End(-4159).Column
    $parts = [Collections.Generic.List[string]]::new()
    for ($row = 1; $row -le $lastRow; $row++) {
        for ($column = 1; $column -le $lastColumn; $column++) { [void]$parts.Add([string]$Worksheet.Cells.Item($row, $column).Value2) }
        [void]$parts.Add("<ROW>")
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($parts -join "`0"))
    $sha = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '') } finally { $sha.Dispose() }
}

function Get-RegistryHashes {
    param([Parameter(Mandatory = $true)][object]$Workbook)
    $result = @{}
    foreach ($name in @('Employees', 'EmployeeCurrentState', 'PersonnelEvents', 'PersonnelStateSnapshots', 'PaymentAssignments', 'DocumentRegistry', 'LegalActs', 'PaymentRules', 'PositionClassification', 'Штат')) {
        try { $result[$name] = Get-SheetHash -Worksheet $Workbook.Worksheets.Item($name) } catch {}
    }
    $result
}

function Assert-NoHashChanges {
    param([hashtable]$Before, [hashtable]$After, [string]$Label)
    foreach ($key in $Before.Keys) {
        if (-not $After.ContainsKey($key) -or $Before[$key] -ne $After[$key]) { throw "$Label mutated sheet $key." }
    }
}

function Release-ComObject {
    param([Parameter(Mandatory = $false)][object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word is running. Close Office before the integrity test.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testDirectory = Join-Path $projectRoot "Trash\personnel-data-integrity-$stamp"
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$cleanPath = Join-Path $testDirectory 'CreateOrder.integrity-clean.xlsm'
$corruptPath = Join-Path $testDirectory 'CreateOrder.integrity-corrupt.xlsm'
$schemaPath = Join-Path $testDirectory 'CreateOrder.integrity-schema.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $cleanPath
Copy-Item -LiteralPath $cleanPath -Destination $corruptPath
Copy-Item -LiteralPath $cleanPath -Destination $schemaPath

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($cleanPath, 0, $false)
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelDataIntegrity' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelDataIntegrity.bas')
    Add-Probe -Workbook $workbook
    Seed-ValidFixture -Workbook $workbook -WorkbookFilePath $cleanPath
    $cleanBefore = Get-RegistryHashes -Workbook $workbook
    $cleanResult = [string]$excel.Run("'$($workbook.Name)'!personnel_integrity_probe.RunIntegrityProbe")
    $cleanAfter = Get-RegistryHashes -Workbook $workbook
    Assert-NoHashChanges -Before $cleanBefore -After $cleanAfter -Label 'Clean scan'
    if ($cleanResult -notlike 'OK|0|*') { throw "Clean integrity fixture is not clean: $cleanResult" }
    $workbook.Close($false)
    Release-ComObject -Value $workbook
    $workbook = $null

    $workbook = $excel.Workbooks.Open($corruptPath, 0, $false)
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelDataIntegrity' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelDataIntegrity.bas')
    Add-Probe -Workbook $workbook
    Seed-ValidFixture -Workbook $workbook -WorkbookFilePath $corruptPath
    Add-CorruptFixture -Workbook $workbook -WorkbookFilePath $corruptPath
    $corruptBefore = Get-RegistryHashes -Workbook $workbook
    $corruptResult = [string]$excel.Run("'$($workbook.Name)'!personnel_integrity_probe.RunIntegrityProbe")
    $corruptAfter = Get-RegistryHashes -Workbook $workbook
    Assert-NoHashChanges -Before $corruptBefore -After $corruptAfter -Label 'Corrupt scan'
    if ($corruptResult -notlike 'OK|*|*') { throw "Corrupt integrity probe failed: $corruptResult" }
    $corruptParts = $corruptResult -split '\|', 4
    $categories = ($corruptParts[2] -split ';' | Where-Object { $_ } | ForEach-Object { ($_ -split '=')[0] }) | Sort-Object -Unique
    $requiredCategories = @('IDENTIFIER', 'EVENT_EMPLOYEE', 'EVENT_SNAPSHOTS', 'SNAPSHOT_LINKAGE', 'ASSIGNMENT_LINKAGE', 'DOCUMENT_LINKAGE', 'CURRENT_STATE_LINKAGE', 'CHRONOLOGY', 'ACTIVE_PAYMENTS', 'STAFF_LINK', 'LEGAL_REFERENCE')
    foreach ($category in $requiredCategories) { if ($categories -notcontains $category) { throw "Corrupt fixture did not produce category $category. Result: $corruptResult" } }
    if ($corruptParts[3] -match 'Fixture Employee|PN-1|E-MISSING') { throw 'Integrity report contains fixture PII.' }
    $workbook.Close($false)
    Release-ComObject -Value $workbook
    $workbook = $null

    $workbook = $excel.Workbooks.Open($schemaPath, 0, $false)
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelDataIntegrity' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelDataIntegrity.bas')
    Add-Probe -Workbook $workbook
    Seed-ValidFixture -Workbook $workbook -WorkbookFilePath $schemaPath
    $rules = $workbook.Worksheets.Item('PaymentRules')
    $null = $rules.Cells.Item(1, (Get-HeaderMap -Worksheet $rules)['ActID']).Value2 = ''
    $schemaBefore = Get-RegistryHashes -Workbook $workbook
    $schemaResult = [string]$excel.Run("'$($workbook.Name)'!personnel_integrity_probe.RunIntegrityProbe")
    $schemaAfter = Get-RegistryHashes -Workbook $workbook
    Assert-NoHashChanges -Before $schemaBefore -After $schemaAfter -Label 'Schema scan'
    if ($schemaResult -notlike 'OK|*|SCHEMA=ERROR;*') { throw "Schema fixture did not produce SCHEMA: $schemaResult" }
    Release-ComObject -Value $rules
    $rules = $null
    $workbook.Close($false)
    Release-ComObject -Value $workbook
    $workbook = $null
    $excel.Quit()
    Release-ComObject -Value $excel
    $excel = $null
    Write-Output "PERSONNEL_DATA_INTEGRITY_OK|clean=$cleanResult|corrupt_categories=$($categories -join ',')|schema=$schemaResult"
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject -Value $workbook }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject -Value $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

for ($attempt = 1; $attempt -le 120 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained running after the integrity test.' }
