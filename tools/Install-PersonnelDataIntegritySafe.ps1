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
    if ($component.Type -ne 1) { throw "$ModuleName exists but is not a standard module." }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $null = $module.DeleteLines(1, $module.CountOfLines) }
    $null = $module.AddFromString($code)
}

function Invoke-IntegrityReport {
    param(
        [Parameter(Mandatory = $true)][object]$Excel,
        [Parameter(Mandatory = $true)][object]$Workbook
    )
    [string]$Excel.Run("'$($Workbook.Name)'!mdlPersonnelDataIntegrity.BuildPersonnelDataIntegrityReport")
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word is running. Close Office before the integrity module import.' }
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$modulePath = Join-Path $resolvedSource 'mdlPersonnelDataIntegrity.bas'
if (-not (Test-Path -LiteralPath $modulePath)) { throw "Missing integrity module source: $modulePath" }
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stagingDirectory = Join-Path $projectRoot "Trash\personnel-data-integrity-install-probe-$stamp"
$backupDirectory = Join-Path $projectRoot "CreateOrderBackups\personnel-data-integrity-installed-$stamp"
New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$stagingWorkbook = Join-Path $stagingDirectory 'CreateOrder.integrity-install-probe.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-personnel-data-integrity-install.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $stagingWorkbook

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($stagingWorkbook, 0, $false)
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelDataIntegrity' -ModulePath $modulePath
    $stagingReport = Invoke-IntegrityReport -Excel $excel -Workbook $workbook
    if ($stagingReport -notmatch 'findings=\d+; errors=\d+; warnings=\d+') { throw "Staging integrity report is invalid: $stagingReport" }
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

for ($attempt = 1; $attempt -le 60 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after staging integrity import.' }

Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)
    Import-CodeModuleText -Workbook $workbook -ModuleName 'mdlPersonnelDataIntegrity' -ModulePath $modulePath
    $null = $workbook.Save()
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

for ($attempt = 1; $attempt -le 60 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after working integrity import.' }

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
    $postInstallReport = Invoke-IntegrityReport -Excel $excel -Workbook $workbook
    if ($postInstallReport -notmatch 'findings=\d+; errors=\d+; warnings=\d+') { throw "Post-install integrity report is invalid: $postInstallReport" }
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

for ($attempt = 1; $attempt -le 60 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 500 }
if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Office process remained after post-install integrity verification.' }
Write-Output "PERSONNEL_DATA_INTEGRITY_INSTALL_OK|backup=$backupWorkbook|staging=$stagingWorkbook|$postInstallReport"
