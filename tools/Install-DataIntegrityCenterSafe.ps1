[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$RibbonXmlPath
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }
if ([string]::IsNullOrWhiteSpace($RibbonXmlPath)) { $RibbonXmlPath = Join-Path $PSScriptRoot '..\resources\customUI14.xml' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-VbaText {
    param([Parameter(Mandatory = $true)][string]$Path)
    [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Import-CodeModuleText {
    param([object]$Workbook, [string]$ModuleName, [string]$ModulePath)
    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $null
    try { $component = $Workbook.VBProject.VBComponents.Item($ModuleName) } catch {}
    if ($null -eq $component) { $component = $Workbook.VBProject.VBComponents.Add(1); $component.Name = $ModuleName }
    if ($component.Type -ne 1) { throw "$ModuleName exists but is not a standard module." }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $null = $module.DeleteLines(1, $module.CountOfLines) }
    $null = $module.AddFromString($code)
}

function Import-UserForm {
    param([object]$Workbook, [string]$FormName, [string]$FormPath)
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item($FormName)) } catch {}
    $component = $Workbook.VBProject.VBComponents.Import($FormPath)
    if ($component.Name -ne $FormName -or $component.Type -ne 3) { throw "Unexpected imported form: $($component.Name)/$($component.Type)" }
}

function Ensure-IntegrityLocalization {
    param([object]$Workbook)
    $sheet = $Workbook.Worksheets.Item('Localization')
    $lastColumn = $sheet.Cells.Item(1, $sheet.Columns.Count).End(-4159).Column
    $ruColumn = 0
    for ($column = 2; $column -le $lastColumn; $column++) {
        if ([string]::Equals([string]$sheet.Cells.Item(1, $column).Value2, 'ru', [StringComparison]::OrdinalIgnoreCase)) { $ruColumn = $column; break }
    }
    if ($ruColumn -eq 0) { throw 'Localization sheet does not have a ru column.' }
    $keyRows = @{}
    $lastRow = $sheet.Cells.Find('*', $sheet.Cells.Item(1, 1), -4123, 1, 1, 2, $false).Row
    for ($row = 2; $row -le $lastRow; $row++) {
        $key = ([string]$sheet.Cells.Item($row, 1).Value2).Trim()
        if ($key) { $keyRows[$key.ToLowerInvariant()] = $row }
    }
    $translations = [ordered]@{
        'integrity.form.title' = 'Центр целостности данных'
        'integrity.form.description' = 'Диагностика кадровых реестров только для чтения.'
        'integrity.form.severity' = 'Уровень'
        'integrity.form.category' = 'Категория'
        'integrity.form.readonly' = 'Только чтение: исправление не выполняется.'
        'integrity.form.scan' = 'Проверить'
        'integrity.form.close' = 'Закрыть'
        'integrity.form.all' = 'ВСЕ'
        'integrity.form.not_scanned' = 'Проверка еще не запускалась.'
        'integrity.form.scan_failed' = 'Проверка целостности не выполнена.'
    }
    foreach ($key in $translations.Keys) {
        if ($keyRows.ContainsKey($key.ToLowerInvariant())) { $row = $keyRows[$key.ToLowerInvariant()] } else { $lastRow++; $row = $lastRow; $null = $sheet.Cells.Item($row, 1).Value2 = [string]$key }
        $null = $sheet.Cells.Item($row, $ruColumn).Value2 = [string]$translations[$key]
    }
}

function Set-RibbonXmlInWorkbook {
    param([string]$WorkbookFilePath, [string]$XmlFilePath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ('CreateOrderRibbon-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempDirectory -Force | Out-Null
    $zipPath = Join-Path $tempDirectory 'book.zip'
    Copy-Item -LiteralPath $WorkbookFilePath -Destination $zipPath
    $unpack = Join-Path $tempDirectory 'unpack'
    Expand-Archive -LiteralPath $zipPath -DestinationPath $unpack -Force
    $customUi = Join-Path $unpack 'customUI'
    New-Item -ItemType Directory -Path $customUi -Force | Out-Null
    Copy-Item -LiteralPath $XmlFilePath -Destination (Join-Path $customUi 'customUI14.xml') -Force
    $relsPath = Join-Path $unpack '_rels\.rels'
    [xml]$rels = Get-Content -LiteralPath $relsPath -Raw
    $ns = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
    if ($null -eq $rels.SelectSingleNode("//r:Relationship[@Target='customUI/customUI14.xml']", $ns)) {
        $node = $rels.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')
        $node.SetAttribute('Id', 'customUIRelID')
        $node.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
        $node.SetAttribute('Target', 'customUI/customUI14.xml')
        $null = $rels.DocumentElement.AppendChild($node)
        $rels.Save($relsPath)
    }
    $outputZip = Join-Path $tempDirectory 'output.zip'
    $archive = [IO.Compression.ZipFile]::Open($outputZip, [IO.Compression.ZipArchiveMode]::Create)
    try {
        Get-ChildItem -LiteralPath $unpack -Recurse -File | ForEach-Object {
            $entryName = $_.FullName.Substring($unpack.Length + 1).Replace('\', '/')
            [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $_.FullName, $entryName, [IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
    } finally { $archive.Dispose() }
    Copy-Item -LiteralPath $outputZip -Destination $WorkbookFilePath -Force
    Move-Item -LiteralPath $tempDirectory -Destination (Join-Path ([IO.Path]::GetDirectoryName($WorkbookFilePath)) ('.integrity-ribbon-temp-' + [guid]::NewGuid().ToString('N'))) -Force
    $movedTemp = Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($WorkbookFilePath)) -Directory | Where-Object Name -like '.integrity-ribbon-temp-*' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -ne $movedTemp) { Remove-Item -LiteralPath $movedTemp.FullName -Recurse -Force }
}

function Assert-RibbonButton {
    param([string]$WorkbookFilePath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($WorkbookFilePath)
    try {
        $entry = $archive.GetEntry('customUI/customUI14.xml')
        if ($null -eq $entry) { throw 'Workbook has no customUI14.xml.' }
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
        try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        if ($xml -notmatch 'openDataIntegrityCenter' -or $xml -notmatch 'OnOpenDataIntegrityCenterClick') { throw 'Ribbon XML does not contain the data integrity button.' }
    } finally { $archive.Dispose() }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word is running. Close Office before the integrity center import.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$resolvedRibbon = (Resolve-Path -LiteralPath $RibbonXmlPath).Path
$formPath = Join-Path $resolvedSource 'frmDataIntegrityCenter.frm'
foreach ($path in @((Join-Path $resolvedSource 'mdlPersonnelDataIntegrity.bas'), (Join-Path $resolvedSource 'ModuleLocalization.bas'), (Join-Path $resolvedSource 'mdlRibbonHandlers.bas'), $formPath, (Join-Path $resolvedSource 'frmDataIntegrityCenter.frx'), $resolvedRibbon)) { if (-not (Test-Path -LiteralPath $path)) { throw "Missing P2 source: $path" } }

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stagingDirectory = Join-Path $projectRoot "Trash\data-integrity-center-install-probe-$stamp"
$backupDirectory = Join-Path $projectRoot "CreateOrderBackups\data-integrity-center-installed-$stamp"
New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$stagingWorkbook = Join-Path $stagingDirectory 'CreateOrder.data-integrity-center-install-probe.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-data-integrity-center-install.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $stagingWorkbook

function Install-IntoWorkbook {
    param([string]$TargetPath, [bool]$Save)
    $excel = $null
    $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1
        $book = $excel.Workbooks.Open($TargetPath, 0, $false)
        Import-CodeModuleText -Workbook $book -ModuleName 'mdlPersonnelDataIntegrity' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelDataIntegrity.bas')
        Import-CodeModuleText -Workbook $book -ModuleName 'ModuleLocalization' -ModulePath (Join-Path $resolvedSource 'ModuleLocalization.bas')
        Import-CodeModuleText -Workbook $book -ModuleName 'mdlRibbonHandlers' -ModulePath (Join-Path $resolvedSource 'mdlRibbonHandlers.bas')
        Import-UserForm -Workbook $book -FormName 'frmDataIntegrityCenter' -FormPath $formPath
        Ensure-IntegrityLocalization -Workbook $book
        $report = [string]$excel.Run("'$($book.Name)'!mdlPersonnelDataIntegrity.BuildPersonnelDataIntegrityReport")
        if ($report -notmatch 'findings=\d+; errors=\d+; warnings=\d+') { throw "Integrity report did not run: $report" }
        if ($Save) { $null = $book.Save() }
        $book.Close($false)
        $book = $null
        $excel.Quit()
        $excel = $null
        return $report
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$stagingReport = Install-IntoWorkbook -TargetPath $stagingWorkbook -Save $true
Set-RibbonXmlInWorkbook -WorkbookFilePath $stagingWorkbook -XmlFilePath $resolvedRibbon
Assert-RibbonButton -WorkbookFilePath $stagingWorkbook
for ($attempt = 1; $attempt -le 60 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after staging center import.' }

Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook
$workingReport = Install-IntoWorkbook -TargetPath $resolvedWorkbook -Save $true
Set-RibbonXmlInWorkbook -WorkbookFilePath $resolvedWorkbook -XmlFilePath $resolvedRibbon
Assert-RibbonButton -WorkbookFilePath $resolvedWorkbook
for ($attempt = 1; $attempt -le 60 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after working center import.' }
Write-Output "DATA_INTEGRITY_CENTER_INSTALL_OK|backup=$backupWorkbook|staging=$stagingWorkbook|$workingReport"
