[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmPersonnelActionWizardV2'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

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
        operation = 'Install-PersonnelActionWizardV2Safe'
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

if (-not ('PersonnelV2InstallNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PersonnelV2InstallNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][PersonnelV2InstallNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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

function Install-FormIntoWorkbook {
    param(
        [Parameter(Mandatory = $true)][string]$TargetWorkbook,
        [Parameter(Mandatory = $true)][string]$FormPath,
        [Parameter(Mandatory = $true)][string]$ComponentName
    )

    $excel = $null
    $excelProcessId = 0
    $book = $null
    $components = $null
    $existing = $null
    $imported = $null
    try {
        Write-InstallLog INFO 'Opening workbook for personnel V2 import.' @{ workbook = $TargetWorkbook }
        $excel = New-Object -ComObject Excel.Application
        $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1
        $book = $excel.Workbooks.Open($TargetWorkbook, 0, $false)
        if ($book.ReadOnly) { throw "Workbook opened read-only: $TargetWorkbook" }
        if ($book.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot install the personnel V2 form.' }

        $components = $book.VBProject.VBComponents
        try { $existing = $components.Item($ComponentName) } catch { $existing = $null }
        if ($existing) { $components.Remove($existing) }
        $imported = $components.Import($FormPath)
        if ($imported.Name -ne $ComponentName -or $imported.Type -ne 3) {
            throw "Imported component verification failed: name=$($imported.Name), type=$($imported.Type)"
        }
        $original = $null
        try {
            $original = $components.Item('frmPersonnelActionWizard')
            if ($original.Type -ne 3) { throw 'Original personnel action form has an unexpected type.' }
        } finally {
            Release-ComObject $original
        }
        $book.Save()
        Write-InstallLog INFO 'Personnel V2 form imported and workbook saved.' @{ workbook = $TargetWorkbook; component = $ComponentName }
    } finally {
        if ($book) { $book.Close($false) }
        if ($excel) { $excel.Quit() }
        Release-ComObject $imported
        Release-ComObject $existing
        Release-ComObject $components
        Release-ComObject $book
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
    }
}

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
if (-not $resolvedWorkbook.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook must be inside the CreateOrder project: $resolvedWorkbook"
}
if (-not $resolvedSource.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Source directory must be inside the CreateOrder project: $resolvedSource"
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel is running. Save and close Excel before installing the personnel V2 form.'
}

$formPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$frxPath = Join-Path $resolvedSource ($TargetComponentName + '.frx')
$manifestPath = Join-Path $resolvedSource ($TargetComponentName + '.layout.csv')
foreach ($path in @($formPath, $frxPath, $manifestPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing personnel V2 artifact: $path" }
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stagingDirectory = Join-Path $projectRoot ("Trash\personnel-action-v2-install-probe-$stamp")
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\personnel-action-v2-installed-$stamp")
New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$stagingWorkbook = Join-Path $stagingDirectory 'CreateOrder.personnel-v2-install-probe.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-personnel-action-v2-install.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $stagingWorkbook
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook

Write-InstallLog INFO 'Prepared staging workbook and safety backup.' @{
    stagingWorkbook = $stagingWorkbook
    backupWorkbook = $backupWorkbook
}

Install-FormIntoWorkbook -TargetWorkbook $stagingWorkbook -FormPath $formPath -ComponentName $TargetComponentName
for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 250 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel remained running after staging import.' }

& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardV2Designer.ps1') `
    -WorkbookPath $stagingWorkbook `
    -SourceDirectory $resolvedSource `
    -TargetComponentName $TargetComponentName

Install-FormIntoWorkbook -TargetWorkbook $resolvedWorkbook -FormPath $formPath -ComponentName $TargetComponentName
for ($attempt = 1; $attempt -le 20 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 250 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel remained running after working-book import.' }

& (Join-Path $PSScriptRoot 'Test-PersonnelActionWizardV2Designer.ps1') `
    -WorkbookPath $resolvedWorkbook `
    -SourceDirectory $resolvedSource `
    -TargetComponentName $TargetComponentName

Write-InstallLog INFO 'Personnel action V2 safe installation completed.' @{
    workbook = $resolvedWorkbook
    component = $TargetComponentName
    backup = $backupWorkbook
}
