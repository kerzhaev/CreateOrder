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

if (-not ('EnrollmentV2LogicTestNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EnrollmentV2LogicTestNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][EnrollmentV2LogicTestNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$formPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$frxPath = Join-Path $resolvedSource ($TargetComponentName + '.frx')
$bindingPath = Join-Path $resolvedSource ($TargetComponentName + '.bindings.csv')
foreach ($path in @($formPath, $frxPath, $bindingPath)) {
    Assert-Condition (Test-Path -LiteralPath $path) "Missing V2 logic artifact: $path"
}
Assert-Condition (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) 'Excel must be closed before the isolated V2 logic test.'

$formBytes = [IO.File]::ReadAllBytes($formPath)
$formText = [Text.Encoding]::GetEncoding(1251).GetString($formBytes)
Assert-Condition (-not $formText.Contains('Controls.Add(')) 'V2 must not create controls at runtime.'
Assert-Condition (-not [regex]::IsMatch($formText, '\.(Left|Top|Width|Height)\s*=')) 'V2 must not override designer geometry.'
Assert-Condition ($formText.Contains('Private Sub BindDesignerControls()')) 'V2 is missing design-time control binding.'
Assert-Condition ($formText.Contains('Private Sub ApplyDesignerLocalization()')) 'V2 is missing localization binding.'
Assert-Condition (@(Import-Csv -LiteralPath $bindingPath).Count -ge 160) 'V2 binding manifest is unexpectedly small.'

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testDirectory = Join-Path $projectRoot ("Trash\enrollment-designer-v2-logic-test-$stamp")
New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
$testWorkbook = Join-Path $testDirectory 'CreateOrder.v2-logic-test.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $testWorkbook

$probeCode = @'
Option Explicit

Public Sub RunEnrollmentV2LogicProbe()
    Dim resultText As String
    Dim targetSheet As Worksheet
    On Error GoTo ProbeError

    Load frmEnrollmentWizardV2
    resultText = "OK|" & CStr(frmEnrollmentWizardV2.Controls("mpWizard").Pages.Count) & "|" & _
        CStr(frmEnrollmentWizardV2.Controls("btnCheckDynamic").Caption) & "|" & _
        CStr(frmEnrollmentWizardV2.Controls("btnSaveCardDynamic").Caption) & "|" & _
        CStr(frmEnrollmentWizardV2.Controls("btnExportPackageDynamic").Caption) & "|" & _
        CStr(frmEnrollmentWizardV2.Controls("btnClose").Caption)
    Unload frmEnrollmentWizardV2
    GoTo WriteResult

ProbeError:
    resultText = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
    On Error Resume Next
    Unload frmEnrollmentWizardV2
    On Error GoTo 0

WriteResult:
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("__V2LogicProbe").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set targetSheet = ThisWorkbook.Worksheets.Add
    targetSheet.Name = "__V2LogicProbe"
    targetSheet.Range("A1").Value = resultText
End Sub
'@

$excel = $null
$excelProcessId = 0
$book = $null
$components = $null
$existingForm = $null
$importedForm = $null
$probeComponent = $null
$probeSheet = $null
$resultText = ''
try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($testWorkbook, 0, $false)
    $components = $book.VBProject.VBComponents
    try {
        $existingForm = $components.Item($TargetComponentName)
        $components.Remove($existingForm)
    } catch {
        if ($_.Exception.Message -notmatch 'Subscript out of range|Индекс находится вне границ') { throw }
    } finally {
        Release-ComObject $existingForm
    }
    $importedForm = $components.Import($formPath)
    Assert-Condition ($importedForm.Name -eq $TargetComponentName -and $importedForm.Type -eq 3) 'V2 import verification failed in the isolated workbook.'

    $probeComponent = $components.Add(1)
    $probeComponent.Name = 'modEnrollmentV2LogicProbe'
    $probeComponent.CodeModule.AddFromString($probeCode)
    $excel.Run("'$($book.Name)'!modEnrollmentV2LogicProbe.RunEnrollmentV2LogicProbe")
    $probeSheet = $book.Worksheets.Item('__V2LogicProbe')
    $resultText = [string]$probeSheet.Range('A1').Value2
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $probeSheet
    Release-ComObject $probeComponent
    Release-ComObject $importedForm
    Release-ComObject $components
    Release-ComObject $book
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}

for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) {
    Start-Sleep -Milliseconds 250
}
Assert-Condition (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) 'Excel remained running after the isolated V2 logic test.'
Assert-Condition ($resultText.StartsWith('OK|7|')) "V2 initialization probe failed: $resultText"
Write-Host "Enrollment Wizard V2 logic verification passed: $resultText"
