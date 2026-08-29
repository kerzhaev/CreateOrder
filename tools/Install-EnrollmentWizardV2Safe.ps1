[CmdletBinding()]
param(
    [string]$WorkbookPath
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-InstallLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$Context = @{}
    )
    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = 'Install-EnrollmentWizardV2Safe'
        message = $Message
    }
    foreach ($key in $Context.Keys) { $payload[$key] = $Context[$key] }
    $line = $payload | ConvertTo-Json -Compress -Depth 5
    if ($Level -eq 'DEBUG') { Write-Verbose $line }
    elseif ($Level -eq 'WARN') { Write-Warning $line }
    elseif ($Level -eq 'ERROR') { Write-Error $line }
    else { Write-Host $line }
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

if (-not ('EnrollmentV2InstallNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EnrollmentV2InstallNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][EnrollmentV2InstallNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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
    if ($process -and $process.ProcessName -eq 'EXCEL') {
        Write-InstallLog WARN 'Excel did not exit after COM Quit; stopping only the installer-owned process.' @{ processId = $ProcessId }
        Stop-Process -Id $ProcessId -Force
    }
}

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
    $code = Read-VbaText $ModulePath
    $code = [regex]::Replace($code, '^Attribute VB_Name\s*=\s*"[^"]+"\r?\n', '', 1)
    $component = $null
    $codeModule = $null
    try {
        $component = $Workbook.VBProject.VBComponents.Item($ModuleName)
        $codeModule = $component.CodeModule
        if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
        $codeModule.AddFromString($code)
    } finally {
        Release-ComObject $codeModule
        Release-ComObject $component
    }
}

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
if (-not $resolvedWorkbook.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook must be inside the CreateOrder project: $resolvedWorkbook"
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel is open. Close all Excel windows before installing Enrollment Wizard V2.'
}

Write-InstallLog INFO 'Running isolated V2 preflight before changing the working workbook.' @{ workbook = $resolvedWorkbook }
& (Join-Path $PSScriptRoot 'Test-EnrollmentWizardV2Logic.ps1') -WorkbookPath $resolvedWorkbook
if (-not $?) { throw 'Enrollment Wizard V2 logic preflight failed.' }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\enrollment-designer-v2-installed-$stamp")
$backupPath = Join-Path $backupDirectory 'CreateOrder.before-enrollment-designer-v2-activation.xlsm'
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupPath
Write-InstallLog INFO 'Created the pre-installation workbook backup.' @{ backup = $backupPath }

$moduleDirectory = Join-Path $projectRoot 'CreateOrder.xlsm.modules'
$formPath = Join-Path $moduleDirectory 'frmEnrollmentWizardV2.frm'
$excel = $null
$excelProcessId = 0
$book = $null
$components = $null
$existingForm = $null
$importedForm = $null
$probeComponent = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)
    if ($book.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot install V2.' }

    Import-CodeModuleText -Workbook $book -ModuleName 'mdlEnrollmentWorkflow' -ModulePath (Join-Path $moduleDirectory 'mdlEnrollmentWorkflow.bas')
    Import-CodeModuleText -Workbook $book -ModuleName 'mdlRibbonHandlers' -ModulePath (Join-Path $moduleDirectory 'mdlRibbonHandlers.bas')

    $components = $book.VBProject.VBComponents
    try {
        $existingForm = $components.Item('frmEnrollmentWizardV2')
        $components.Remove($existingForm)
    } catch {
        if ($_.Exception.Message -notmatch 'Subscript out of range|Индекс находится вне границ') { throw }
    } finally {
        Release-ComObject $existingForm
    }
    $importedForm = $components.Import($formPath)
    if ($importedForm.Name -ne 'frmEnrollmentWizardV2' -or $importedForm.Type -ne 3) {
        throw "V2 form import verification failed: name=$($importedForm.Name), type=$($importedForm.Type)"
    }
    if ($components.Item('frmEnrollmentWizard').Type -ne 3) { throw 'V1 backup form is missing after V2 installation.' }

    $probeComponent = $components.Add(1)
    $probeComponent.Name = 'modEnrollmentV2InstallProbe'
    $probeComponent.CodeModule.AddFromString(@'
Option Explicit
Public Sub VerifyEnrollmentV2Install()
    Load frmEnrollmentWizardV2
    If frmEnrollmentWizardV2.Controls("mpWizard").Pages.Count <> 7 Then Err.Raise 5, , "V2 page count mismatch"
    Unload frmEnrollmentWizardV2
End Sub
'@)
    $excel.Run("'$($book.Name)'!modEnrollmentV2InstallProbe.VerifyEnrollmentV2Install")
    $components.Remove($probeComponent)
    Release-ComObject $probeComponent
    $probeComponent = $null

    $book.Save()
    Write-InstallLog INFO 'Installed and initialized V2; active modules now route enrollment to it.' @{
        workbook = $resolvedWorkbook
        backup = $backupPath
        retainedFallback = 'frmEnrollmentWizard'
    }
} catch {
    Write-InstallLog ERROR 'V2 installation failed; the working workbook backup is available.' @{ error = $_.Exception.Message; backup = $backupPath }
    throw
} finally {
    if ($probeComponent -and $components) { try { $components.Remove($probeComponent) } catch {} }
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
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
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel remained running after V2 installation.' }

Write-InstallLog INFO 'Enrollment Wizard V2 installation completed.' @{ workbook = $resolvedWorkbook; backup = $backupPath }
