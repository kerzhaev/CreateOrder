[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmEnrollmentWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

if (-not ('EnrollmentDesignerTestNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EnrollmentDesignerTestNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][EnrollmentDesignerTestNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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

function Assert-Condition {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) { throw $Message }
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$frmPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$frxPath = Join-Path $resolvedSource ($TargetComponentName + '.frx')
$manifestPath = Join-Path $resolvedSource ($TargetComponentName + '.layout.csv')

foreach ($path in @($frmPath, $frxPath, $manifestPath)) {
    Assert-Condition (Test-Path -LiteralPath $path) "Missing designer artifact: $path"
}
Assert-Condition (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) 'Excel must be closed before the V2 designer verification.'

$formBytes = [IO.File]::ReadAllBytes($frmPath)
$formText = [Text.Encoding]::GetEncoding(1251).GetString($formBytes)
$lfCount = 0
$crlfCount = 0
for ($index = 0; $index -lt $formBytes.Length; $index++) {
    if ($formBytes[$index] -eq 10) {
        $lfCount++
        if ($index -gt 0 -and $formBytes[$index - 1] -eq 13) { $crlfCount++ }
    }
}
Assert-Condition ($lfCount -eq $crlfCount) 'The exported .frm must use CRLF line endings.'
Assert-Condition ($formText.Contains('Attribute VB_Name = "frmEnrollmentWizardV2"')) 'The exported .frm has an unexpected VB_Name.'
Assert-Condition ($formText.Contains('Private Sub BindDesignerControls()')) 'The V2 form is missing design-time control binding.'
Assert-Condition ($formText.Contains('Private Sub ApplyDesignerLocalization()')) 'The V2 form is missing designer localization binding.'
Assert-Condition (-not $formText.Contains('Controls.Add(')) 'The V2 form must not create runtime controls.'
Assert-Condition (-not [regex]::IsMatch($formText, '\.(Left|Top|Width|Height)\s*=')) 'The V2 form code must not override owner geometry.'

$manifest = @(Import-Csv -LiteralPath $manifestPath)
$expectedPages = @('pgEmployee', 'pgDocs', 'pgMonthly', 'pgOneTime', 'pgAdvanced', 'pgExtras', 'pgPreview')
$pages = @($manifest | Where-Object control_type -eq 'Page')
$designerNames = @($manifest | ForEach-Object designer_name)
$uniqueDesignerNames = @($designerNames | Sort-Object -Unique)
Assert-Condition ($manifest.Count -ge 250) "Layout manifest is unexpectedly small: $($manifest.Count) rows."
Assert-Condition ($pages.Count -eq 7) "Expected seven pages in the manifest; found $($pages.Count)."
Assert-Condition ($designerNames.Count -eq $uniqueDesignerNames.Count) 'Designer control names must be globally unique.'
foreach ($pageName in $expectedPages) {
    Assert-Condition ($pageName -in $pages.designer_name) "Missing page in manifest: $pageName"
}
Assert-Condition (@($manifest | Where-Object { $_.designer_name -eq 'mpWizard' -and $_.control_type -eq 'MultiPage' }).Count -eq 1) 'Manifest must contain one mpWizard MultiPage.'
Assert-Condition (@($manifest | Where-Object designer_name -eq 'fraOrder727').Count -eq 1) 'Manifest is missing fraOrder727.'
Assert-Condition (@($manifest | Where-Object designer_name -eq 'fraOrder430').Count -eq 1) 'Manifest is missing fraOrder430.'

$excel = $null
$excelProcessId = 0
$book = $null
$v1 = $null
$v2 = $null
$designer = $null
$multiPage = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)

    $v1 = $book.VBProject.VBComponents.Item('frmEnrollmentWizard')
    $v2 = $book.VBProject.VBComponents.Item($TargetComponentName)
    Assert-Condition ($v1.Type -eq 3) "Original form has unexpected component type: $($v1.Type)"
    Assert-Condition ($v2.Type -eq 3) "V2 has unexpected component type: $($v2.Type)"
    $designer = $v2.Designer
    Assert-Condition ([double]$designer.InsideWidth -ge 850) "V2 form client width is unexpectedly small: $($designer.InsideWidth)"
    Assert-Condition ([double]$designer.InsideHeight -ge 680) "V2 form client height is unexpectedly small: $($designer.InsideHeight)"
    Assert-Condition ([string]$designer.Caption -ne 'UserForm1') 'V2 form caption was not copied from the current wizard.'

    $v2Code = $v2.CodeModule.Lines(1, $v2.CodeModule.CountOfLines)
    Assert-Condition ($v2Code.Contains('Private Sub BindDesignerControls()')) 'Installed V2 is not connected to design-time controls.'
    Assert-Condition ($v2Code.Contains('Private Sub ApplyDesignerLocalization()')) 'Installed V2 is missing designer localization.'
    Assert-Condition (-not $v2Code.Contains('CreateWizardUi')) 'Installed V2 unexpectedly contains runtime UI construction.'

    $multiPage = $designer.Controls.Item('mpWizard')
    Assert-Condition ($multiPage.Pages.Count -eq 7) "Installed V2 must contain seven pages; found $($multiPage.Pages.Count)."
    for ($pageIndex = 0; $pageIndex -lt $expectedPages.Count; $pageIndex++) {
        $page = $null
        try {
            $page = $multiPage.Pages.Item($pageIndex)
            Assert-Condition ($page.Name -eq $expectedPages[$pageIndex]) "Unexpected page name at index $pageIndex`: $($page.Name)"
            Assert-Condition ($page.Controls.Count -gt 0) "Designer page has no controls: $($page.Name)"
        } finally {
            Release-ComObject $page
        }
    }

    foreach ($componentName in @('mdlEnrollmentWorkflow', 'mdlRibbonHandlers')) {
        $routingComponent = $null
        try {
            $routingComponent = $book.VBProject.VBComponents.Item($componentName)
            $routingCode = $routingComponent.CodeModule.Lines(1, $routingComponent.CodeModule.CountOfLines)
            Assert-Condition ($routingCode.Contains($TargetComponentName)) "Installed routing module does not reference V2: $componentName"
            Assert-Condition (-not [regex]::IsMatch($routingCode, 'frmEnrollmentWizard(?!V2)')) "Installed routing module still references V1: $componentName"
        } finally {
            Release-ComObject $routingComponent
        }
    }
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $multiPage
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

Assert-Condition (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) 'Excel remained running after V2 designer verification.'

Write-Host ("Enrollment Wizard V2 designer verification passed: {0} manifest rows, 7 pages, V1 retained, V2 active." -f $manifest.Count)
