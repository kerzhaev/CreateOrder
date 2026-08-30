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

function Read-VbaText([string]$Path) { [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251)) }

function Import-CodeModuleText([object]$Workbook, [string]$ModuleName, [string]$ModulePath) {
    $code = [regex]::Replace((Read-VbaText $ModulePath), '(?m)^Attribute .*(?:\r?\n|$)', '')
    $component = $null
    try { $component = $Workbook.VBProject.VBComponents.Item($ModuleName) } catch {}
    if ($null -eq $component) { $component = $Workbook.VBProject.VBComponents.Add(1); $component.Name = $ModuleName }
    if ($component.Type -ne 1) { throw "$ModuleName exists but is not a standard module." }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { [void]$module.DeleteLines(1, $module.CountOfLines) }
    [void]$module.AddFromString($code)
}

function Import-UserForm([object]$Workbook, [string]$FormName, [string]$FormPath) {
    try { $Workbook.VBProject.VBComponents.Remove($Workbook.VBProject.VBComponents.Item($FormName)) } catch {}
    $component = $Workbook.VBProject.VBComponents.Import($FormPath)
    if ($component.Name -ne $FormName -or $component.Type -ne 3) { throw "Unexpected imported form: $($component.Name)/$($component.Type)" }
}

function Ensure-GroupedLocalization([object]$Workbook) {
    $sheet = $Workbook.Worksheets.Item('Localization')
    $lastColumn = $sheet.Cells.Item(1, $sheet.Columns.Count).End(-4159).Column
    $ruColumn = 0
    for ($column = 2; $column -le $lastColumn; $column++) { if ([string]::Equals([string]$sheet.Cells.Item(1, $column).Value2, 'ru', [StringComparison]::OrdinalIgnoreCase)) { $ruColumn = $column; break } }
    if ($ruColumn -eq 0) { throw 'Localization sheet does not have a ru column.' }
    $lastFound = $sheet.Cells.Find('*', $sheet.Cells.Item(1, 1), -4123, 1, 1, 2, $false)
    $lastRow = if ($null -eq $lastFound) { 1 } else { [int]$lastFound.Row }
    $keyRows = @{}
    for ($row = 2; $row -le $lastRow; $row++) { $key = ([string]$sheet.Cells.Item($row, 1).Value2).Trim(); if ($key) { $keyRows[$key.ToLowerInvariant()] = $row } }
    $translations = [ordered]@{
        'personnel.grouped.title' = 'Единый кадровый приказ'
        'personnel.grouped.description' = 'Выберите сохранённые EventID через запятую. Порядок строк и параграфов будет сохранён.'
        'personnel.grouped.selection' = 'EventID (необязательно)'
        'personnel.grouped.read_only' = 'Проверка не изменяет реестры; экспорт блокируется при неполной записи.'
        'personnel.grouped.preview' = 'Проверить'
        'personnel.grouped.export' = 'Сформировать DOCX'
        'personnel.grouped.not_loaded' = 'Проверка ещё не запускалась.'
        'personnel.grouped.valid' = 'Данные готовы к формированию DOCX.'
        'personnel.grouped.invalid' = 'Есть ошибки. Исправьте строки, отмеченные в отчёте.'
        'personnel.grouped.failed' = 'Операция не выполнена: {error}'
        'personnel.grouped.exported' = 'DOCX сформирован: {path}'
    }
    foreach ($key in $translations.Keys) {
        if ($keyRows.ContainsKey($key.ToLowerInvariant())) { $row = $keyRows[$key.ToLowerInvariant()] } else { $lastRow++; $row = $lastRow; $null = $sheet.Cells.Item($row, 1).Value2 = [string]$key }
        $null = $sheet.Cells.Item($row, $ruColumn).Value2 = [string]$translations[$key]
    }
}

function Set-RibbonXmlInWorkbook([string]$WorkbookFilePath, [string]$XmlFilePath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ('CreateOrderGroupedRibbon-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempDirectory -Force | Out-Null
    try {
        $zipPath = Join-Path $tempDirectory 'book.zip'; Copy-Item -LiteralPath $WorkbookFilePath -Destination $zipPath -Force
        $unpack = Join-Path $tempDirectory 'unpack'; Expand-Archive -LiteralPath $zipPath -DestinationPath $unpack -Force
        $customUi = Join-Path $unpack 'customUI'; New-Item -ItemType Directory -Path $customUi -Force | Out-Null
        Copy-Item -LiteralPath $XmlFilePath -Destination (Join-Path $customUi 'customUI14.xml') -Force
        $relsPath = Join-Path $unpack '_rels\.rels'; [xml]$rels = Get-Content -LiteralPath $relsPath -Raw
        $ns = New-Object System.Xml.XmlNamespaceManager($rels.NameTable); $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
        if ($null -eq $rels.SelectSingleNode("//r:Relationship[@Target='customUI/customUI14.xml']", $ns)) {
            $node = $rels.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships'); $node.SetAttribute('Id', 'customUIRelID'); $node.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'); $node.SetAttribute('Target', 'customUI/customUI14.xml'); [void]$rels.DocumentElement.AppendChild($node); $rels.Save($relsPath)
        }
        $outputZip = Join-Path $tempDirectory 'output.zip'; $archive = [IO.Compression.ZipFile]::Open($outputZip, [IO.Compression.ZipArchiveMode]::Create)
        try { Get-ChildItem -LiteralPath $unpack -Recurse -File | ForEach-Object { $entryName = $_.FullName.Substring($unpack.Length + 1).Replace('\', '/'); [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $_.FullName, $entryName, [IO.Compression.CompressionLevel]::Optimal) | Out-Null } } finally { $archive.Dispose() }
        Copy-Item -LiteralPath $outputZip -Destination $WorkbookFilePath -Force
    } finally { Remove-Item -LiteralPath $tempDirectory -Recurse -Force -ErrorAction SilentlyContinue }
}

function Assert-GroupedRibbon([string]$WorkbookFilePath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($WorkbookFilePath)
    try {
        $entry = $archive.GetEntry('customUI/customUI14.xml'); if ($null -eq $entry) { throw 'Workbook has no customUI14.xml.' }
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8); try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        if ($xml -notmatch 'openGroupedPersonnelOrder' -or $xml -notmatch 'OnOpenGroupedPersonnelOrderClick') { throw 'Ribbon XML does not contain the grouped personnel-order button.' }
    } finally { $archive.Dispose() }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word is running. Close Office before the grouped-order import.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$resolvedRibbon = (Resolve-Path -LiteralPath $RibbonXmlPath).Path
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$formPath = Join-Path $resolvedSource 'frmGroupedPersonnelOrderWizard.frm'
foreach ($path in @((Join-Path $resolvedSource 'mdlGroupedPersonnelOrderExport.bas'), (Join-Path $resolvedSource 'ModuleLocalization.bas'), (Join-Path $resolvedSource 'mdlRibbonHandlers.bas'), $formPath, (Join-Path $resolvedSource 'frmGroupedPersonnelOrderWizard.frx'), $resolvedRibbon)) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing P5 source: $path" } }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stagingDirectory = Join-Path $projectRoot "Trash\grouped-personnel-order-install-probe-$stamp"
$backupDirectory = Join-Path $projectRoot "CreateOrderBackups\grouped-personnel-order-installed-$stamp"
New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null; New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$stagingWorkbook = Join-Path $stagingDirectory 'CreateOrder.grouped-install-probe.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-grouped-personnel-order-install.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $stagingWorkbook -Force

function Install-IntoWorkbook([string]$TargetPath, [bool]$Save) {
    $excel = $null; $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 1
        $book = $excel.Workbooks.Open($TargetPath, 0, $false)
        Import-CodeModuleText $book 'mdlGroupedPersonnelOrderExport' (Join-Path $resolvedSource 'mdlGroupedPersonnelOrderExport.bas')
        Import-CodeModuleText $book 'ModuleLocalization' (Join-Path $resolvedSource 'ModuleLocalization.bas')
        Import-CodeModuleText $book 'mdlRibbonHandlers' (Join-Path $resolvedSource 'mdlRibbonHandlers.bas')
        Import-UserForm $book 'frmGroupedPersonnelOrderWizard' $formPath
        Ensure-GroupedLocalization $book
        $report = [string]$excel.Run("'$($book.Name)'!mdlGroupedPersonnelOrderExport.BuildGroupedPersonnelOrderReport")
        if ($report -notmatch '^(OK|INVALID)\|') { throw "Grouped-order module did not run: $report" }
        if ($Save) { [void]$book.Save() }
        $book.Close($false); $book = $null; $excel.Quit(); $excel = $null
        return $report
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

$stagingReport = Install-IntoWorkbook $stagingWorkbook $true
Set-RibbonXmlInWorkbook $stagingWorkbook $resolvedRibbon
Assert-GroupedRibbon $stagingWorkbook
for ($attempt = 1; $attempt -le 120 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after staging grouped-order import.' }

Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook -Force
$workingReport = Install-IntoWorkbook $resolvedWorkbook $true
Set-RibbonXmlInWorkbook $resolvedWorkbook $resolvedRibbon
Assert-GroupedRibbon $resolvedWorkbook
for ($attempt = 1; $attempt -le 120 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after working grouped-order import.' }
Write-Output "GROUPED_PERSONNEL_ORDER_INSTALL_OK|backup=$backupWorkbook|staging=$stagingWorkbook|stagingReport=$stagingReport|workingReport=$workingReport"
