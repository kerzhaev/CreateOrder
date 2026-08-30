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
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1251))
}

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Wait-OfficeExit {
    for ($attempt = 1; $attempt -le 40 -and (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue); $attempt++) {
        Start-Sleep -Milliseconds 250
    }
    if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
        throw 'Office process remained running after the installer-owned operation.'
    }
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
    if ($component.Type -ne 1) { throw "Preview component is not a standard module: $ModuleName" }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($code)
}

function Install-PreviewModule {
    param(
        [Parameter(Mandatory = $true)][string]$TargetWorkbook,
        [Parameter(Mandatory = $true)][string]$ModulePath,
        [Parameter(Mandatory = $true)][string]$OrderTextModulePath,
        [Parameter(Mandatory = $true)][string]$OrderExportModulePath
    )
    $excel = $null
    $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AutomationSecurity = 1 } catch {}
        $book = $excel.Workbooks.Open($TargetWorkbook, 0, $false)
        if ($book.ReadOnly) { throw "Workbook opened read-only: $TargetWorkbook" }
        if ($book.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot install preview module.' }
        Import-CodeModuleText -Workbook $book -ModuleName 'mdlPersonnelOrderText' -ModulePath $OrderTextModulePath
        Import-CodeModuleText -Workbook $book -ModuleName 'mdlPersonnelEventOrderExport' -ModulePath $OrderExportModulePath
        Import-CodeModuleText -Workbook $book -ModuleName 'mdlPersonnelActionPreview' -ModulePath $ModulePath
        $book.Save()
        $component = $book.VBProject.VBComponents.Item('mdlPersonnelActionPreview')
        if ($component.Type -ne 1 -or $component.CodeModule.CountOfLines -lt 10) { throw 'Installed preview module verification failed.' }
    }
    finally {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        Release-ComObject $book
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-InstalledSmoke {
    param([Parameter(Mandatory = $true)][string]$TargetWorkbook)
    $probeCode = @'
Option Explicit
Public Function RunInstalledPreviewSmoke() As String
    Dim draft As Object
    Dim preview As Object
    On Error GoTo Failed
    Set draft = CreateObject("Scripting.Dictionary")
    draft.Add "event_type", "UNKNOWN"
    Set preview = mdlPersonnelActionPreview.BuildPersonnelActionPreview(draft)
    If preview("can_confirm") Then Err.Raise 980, , "Invalid installed preview unexpectedly became confirmable"
    If preview("warnings").Count = 0 Then Err.Raise 981, , "Installed preview did not return a warning"
    RunInstalledPreviewSmoke = "PERSONNEL_ACTION_PREVIEW_INSTALLED_OK"
    Exit Function
Failed:
    RunInstalledPreviewSmoke = "FAILED: " & Err.Description
End Function
'@
    $excel = $null
    $book = $null
    $probe = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AutomationSecurity = 1 } catch {}
        $book = $excel.Workbooks.Open($TargetWorkbook, 0, $false)
        $component = $book.VBProject.VBComponents.Item('mdlPersonnelActionPreview')
        if ($component.Type -ne 1) { throw 'Installed preview component has unexpected type.' }
        $probe = $book.VBProject.VBComponents.Add(1)
        $probe.Name = 'personnel_preview_install_probe'
        $probe.CodeModule.AddFromString($probeCode)
        $result = [string]$excel.Run("'$($book.Name)'!personnel_preview_install_probe.RunInstalledPreviewSmoke")
        if ($result -ne 'PERSONNEL_ACTION_PREVIEW_INSTALLED_OK') { throw $result }
        $book.Close($false)
        $book = $null
        $excel.Quit()
        $excel = $null
        return $result
    }
    finally {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        Release-ComObject $probe
        Release-ComObject $book
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) {
    throw 'Excel or Word is running. Close Office applications before installing the preview module.'
}

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = [IO.Path]::GetFullPath($SourceDirectory)
$modulePath = Join-Path $resolvedSource 'mdlPersonnelActionPreview.bas'
$orderTextModulePath = Join-Path $resolvedSource 'mdlPersonnelOrderText.bas'
$orderExportModulePath = Join-Path $resolvedSource 'mdlPersonnelEventOrderExport.bas'
foreach ($requiredPath in @($modulePath, $orderTextModulePath, $orderExportModulePath)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) { throw "Missing personnel P1.3 source: $requiredPath" }
}
if (-not $resolvedWorkbook.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw "Workbook must be inside the CreateOrder project: $resolvedWorkbook" }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stagingDirectory = Join-Path $projectRoot ("Trash\personnel-action-preview-install-probe-$stamp")
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\personnel-action-preview-installed-$stamp")
New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$stagingWorkbook = Join-Path $stagingDirectory 'CreateOrder.personnel-action-preview-install-probe.xlsm'
$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.before-personnel-action-preview-install.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $stagingWorkbook -Force
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook -Force

Write-Output "Prepared staging workbook: $stagingWorkbook"
Write-Output "Prepared backup: $backupWorkbook"
Install-PreviewModule -TargetWorkbook $stagingWorkbook -ModulePath $modulePath -OrderTextModulePath $orderTextModulePath -OrderExportModulePath $orderExportModulePath
if ((Invoke-InstalledSmoke -TargetWorkbook $stagingWorkbook) -ne 'PERSONNEL_ACTION_PREVIEW_INSTALLED_OK') { throw 'Staging installed smoke failed.' }
Wait-OfficeExit
Write-Output 'Staging import and smoke test passed.'

if (Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue) { throw 'Excel or Word started before working-book import.' }
Install-PreviewModule -TargetWorkbook $resolvedWorkbook -ModulePath $modulePath -OrderTextModulePath $orderTextModulePath -OrderExportModulePath $orderExportModulePath
if ((Invoke-InstalledSmoke -TargetWorkbook $resolvedWorkbook) -ne 'PERSONNEL_ACTION_PREVIEW_INSTALLED_OK') { throw 'Working-book installed smoke failed.' }
Wait-OfficeExit
Write-Output "PERSONNEL_ACTION_PREVIEW_INSTALLED_OK|$resolvedWorkbook|$backupWorkbook"
