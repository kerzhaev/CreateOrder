[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmDataIntegrityCenter'
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
    param([object]$Workbook, [string]$ModuleName, [string]$ModulePath)
    $code = Read-VbaText -Path $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $null
    try { $component = $Workbook.VBProject.VBComponents.Item($ModuleName) } catch {}
    if ($null -eq $component) { $component = $Workbook.VBProject.VBComponents.Add(1); $component.Name = $ModuleName }
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

function Assert-Condition {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel is running. Close Excel before the data integrity designer test.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$formPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$frxPath = Join-Path $resolvedSource ($TargetComponentName + '.frx')
$manifestPath = Join-Path $resolvedSource ($TargetComponentName + '.layout.csv')
foreach ($path in @($formPath, $frxPath, $manifestPath)) { Assert-Condition (Test-Path -LiteralPath $path) "Missing data integrity artifact: $path" }

$formText = Read-VbaText -Path $formPath
Assert-Condition ($formText.Contains(('Attribute VB_Name = "{0}"' -f $TargetComponentName))) 'Unexpected form VB_Name.'
foreach ($requiredCode in @('Private Sub btnScan_Click()', 'Private Sub cboSeverity_Change()', 'Private Sub cboCategory_Change()', 'Private Sub btnClose_Click()', 'Read-only')) { Assert-Condition ($formText.Contains($requiredCode)) "Form is missing $requiredCode." }
Assert-Condition (-not $formText.Contains('Controls.Add(')) 'Data integrity form must not create runtime controls.'
Assert-Condition (-not [regex]::IsMatch($formText, '\.(Left|Top|Width|Height)\s*=')) 'Data integrity form must not override owner geometry at runtime.'
Assert-Condition (-not [regex]::IsMatch($formText, '(?<!\r)\n')) 'Data integrity .frm must use CRLF.'
$manifest = @(Import-Csv -LiteralPath $manifestPath)
Assert-Condition ($manifest.Count -eq 11) "Unexpected data integrity manifest size: $($manifest.Count)."
Assert-Condition (@($manifest.designer_name | Sort-Object -Unique).Count -eq 11) 'Data integrity control names must be unique.'

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testDirectory = Join-Path $projectRoot "Trash\data-integrity-designer-test-$stamp"
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$testWorkbookPath = Join-Path $testDirectory 'CreateOrder.data-integrity-designer-test.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $testWorkbookPath
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($testWorkbookPath, 0, $false)
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelDataIntegrity' -ModulePath (Join-Path $resolvedSource 'mdlPersonnelDataIntegrity.bas')
    Import-UserForm -Workbook $workbook -FormName $TargetComponentName -FormPath $formPath
    $component = $workbook.VBProject.VBComponents.Item($TargetComponentName)
    Assert-Condition ($component.Type -eq 3) 'Data integrity component is not a UserForm.'
    $designer = $component.Designer
    foreach ($name in @('lblTitle', 'lblDescription', 'lblSeverity', 'cboSeverity', 'lblCategory', 'cboCategory', 'lblSummary', 'txtFindings', 'lblReadOnly', 'btnScan', 'btnClose')) {
        try { $null = $designer.Controls.Item($name) } catch { throw "Missing data integrity control: $name" }
    }
    $findings = $designer.Controls.Item('txtFindings')
    Assert-Condition ([bool]$findings.Locked) 'Findings textbox must be locked.'
    Assert-Condition ([bool]$findings.MultiLine) 'Findings textbox must be multiline.'
    Assert-Condition ([int]$findings.ScrollBars -gt 0) 'Findings textbox must have scrollbars.'
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
for ($attempt = 1; $attempt -le 60 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Office process remained after the data integrity designer test.' }
Write-Output "DATA_INTEGRITY_DESIGNER_TEST_OK|form=$formPath|manifest=$manifestPath"
