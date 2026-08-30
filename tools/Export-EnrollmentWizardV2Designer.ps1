[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$SourceDirectory,
    [string]$TargetComponentName = 'frmEnrollmentWizardV2',
    [string]$OperationName = 'Export-EnrollmentWizardV2Designer',
    [string]$BackupPrefix = 'enrollment-designer-v2-owner-layout'
)

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $PSScriptRoot '..\CreateOrder.xlsm' }
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $PSScriptRoot '..\CreateOrder.xlsm.modules' }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic

function Write-ExportLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$Context = @{}
    )

    $payload = [ordered]@{
        timestamp = (Get-Date).ToString('o')
        level = $Level
        operation = $OperationName
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

if (-not ('EnrollmentDesignerExportNativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EnrollmentDesignerExportNativeMethods {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)
    [uint32]$processId = 0
    [void][EnrollmentDesignerExportNativeMethods]::GetWindowThreadProcessId([IntPtr]$ExcelApplication.Hwnd, [ref]$processId)
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
        Write-ExportLog WARN 'Excel did not exit after COM Quit; stopping only the process created by this export.' @{ processId = $ProcessId }
        Stop-Process -Id $ProcessId -Force
    }
}

function Get-ContainerByPath {
    param(
        [Parameter(Mandatory = $true)][object]$Designer,
        [Parameter(Mandatory = $true)][string]$ContainerPath
    )

    $tokens = @($ContainerPath -split '/')
    if ($tokens.Count -eq 0 -or $tokens[0] -ne 'root') { throw "Invalid container path: $ContainerPath" }
    $current = $Designer
    for ($index = 1; $index -lt $tokens.Count; $index++) {
        $token = $tokens[$index]
        $next = $null
        try {
            $next = $current.Controls.Item($token)
        } catch {
            try {
                $next = $current.Pages.Item($token)
            } catch {
                throw "Container token '$token' was not found while resolving '$ContainerPath'."
            }
        }
        $current = $next
    }
    return $current
}

function Get-ControlTypeName {
    param([Parameter(Mandatory = $true)][object]$Control)
    $typeName = [Microsoft.VisualBasic.Information]::TypeName($Control)
    $typeMap = @{
        IMdcText = 'TextBox'
        IMdcList = 'ListBox'
        IMdcCombo = 'ComboBox'
        IMdcCheckBox = 'CheckBox'
        IMdcLabel = 'Label'
        IMdcCommandButton = 'CommandButton'
        ICommandButton = 'CommandButton'
        ILabelControl = 'Label'
        IOptionFrame = 'Frame'
        IMultiPage = 'MultiPage'
        IPage = 'Page'
        IMdcFrame = 'Frame'
        IMdcMultiPage = 'MultiPage'
        IMdcPage = 'Page'
        PageClass = 'Page'
        FrameClass = 'Frame'
        MultiPageClass = 'MultiPage'
        TextBoxClass = 'TextBox'
        LabelClass = 'Label'
        CommandButtonClass = 'CommandButton'
        ComboBoxClass = 'ComboBox'
        CheckBoxClass = 'CheckBox'
        ListBoxClass = 'ListBox'
    }
    if ($typeMap.ContainsKey($typeName)) { return $typeMap[$typeName] }
    return $typeName
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
    throw 'Excel is running. Save the workbook and close Excel before exporting the owner-edited form.'
}

$frmPath = Join-Path $resolvedSource ($TargetComponentName + '.frm')
$frxPath = Join-Path $resolvedSource ($TargetComponentName + '.frx')
$manifestPath = Join-Path $resolvedSource ($TargetComponentName + '.layout.csv')
foreach ($path in @($frmPath, $frxPath, $manifestPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing baseline artifact: $path" }
}

$baselineManifest = @(Import-Csv -LiteralPath $manifestPath)
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupDirectory = Join-Path $projectRoot ("CreateOrderBackups\$BackupPrefix-$stamp")
$stagingDirectory = Join-Path $projectRoot ("Trash\enrollment-designer-v2-export-$stamp")
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null

$backupWorkbook = Join-Path $backupDirectory 'CreateOrder.owner-layout.xlsm'
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupWorkbook
Copy-Item -LiteralPath $frmPath -Destination (Join-Path $backupDirectory ([IO.Path]::GetFileName($frmPath)))
Copy-Item -LiteralPath $frxPath -Destination (Join-Path $backupDirectory ([IO.Path]::GetFileName($frxPath)))
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $backupDirectory ([IO.Path]::GetFileName($manifestPath)))

Write-ExportLog INFO 'Created a safety copy of the owner-edited workbook and baseline artifacts.' @{
    backupDirectory = $backupDirectory
    workbook = $backupWorkbook
    baselineRows = $baselineManifest.Count
}

$stagedFrm = Join-Path $stagingDirectory ($TargetComponentName + '.frm')
$stagedFrx = Join-Path $stagingDirectory ($TargetComponentName + '.frx')
$excel = $null
$excelProcessId = 0
$book = $null
$component = $null
$designer = $null
$rows = @()
$geometryChanges = 0
try {
    Write-ExportLog INFO 'Opening the workbook read-only and reading the edited designer geometry.' @{ workbook = $resolvedWorkbook }
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
    if ($book.VBProject.Protection -ne 0) { throw 'VBA project is protected; cannot export the designer form.' }

    $component = $book.VBProject.VBComponents.Item($TargetComponentName)
    if ($component.Type -ne 3) { throw "Unexpected V2 component type: $($component.Type)" }
    $designer = $component.Designer
    $component.Export($stagedFrm)
    if (-not (Test-Path -LiteralPath $stagedFrm) -or -not (Test-Path -LiteralPath $stagedFrx)) {
        throw 'Excel did not export the expected .frm/.frx pair.'
    }

    foreach ($expected in $baselineManifest) {
        $container = $null
        $control = $null
        try {
            $container = Get-ContainerByPath -Designer $designer -ContainerPath $expected.container_path
            if ($expected.control_type -eq 'Page') {
                $control = $container.Pages.Item($expected.designer_name)
            } else {
                $control = $container.Controls.Item($expected.designer_name)
            }
            $actualType = Get-ControlTypeName -Control $control
            if ($actualType -ne $expected.control_type) {
                throw "Control type changed for $($expected.designer_name): expected $($expected.control_type), actual $actualType"
            }

            $caption = ''
            $left = $null
            $top = $null
            $width = $null
            $height = $null
            try { $caption = [string]$control.Caption } catch {}
            if ($expected.control_type -ne 'Page') {
                $left = [double]$control.Left
                $top = [double]$control.Top
                $width = [double]$control.Width
                $height = [double]$control.Height
                if ([double]$expected.left -ne $left -or [double]$expected.top -ne $top -or
                    [double]$expected.width -ne $width -or [double]$expected.height -ne $height) {
                    $geometryChanges++
                }
            }
            $rows += [pscustomobject]@{
                container_path = $expected.container_path
                source_name = $expected.source_name
                designer_name = $expected.designer_name
                control_type = $expected.control_type
                caption = $caption
                left = $left
                top = $top
                width = $width
                height = $height
                visible = [bool]$control.Visible
                enabled = [bool]$control.Enabled
            }
        } finally {
            Release-ComObject $control
            if ($container -ne $designer) { Release-ComObject $container }
        }
    }

    if ($rows.Count -ne $baselineManifest.Count) {
        throw "Designer structure count changed: expected $($baselineManifest.Count), exported $($rows.Count)"
    }
} catch {
    Write-ExportLog ERROR 'Owner layout export failed; baseline artifacts were not overwritten.' @{
        error = $_.Exception.Message
        backupDirectory = $backupDirectory
        stagingDirectory = $stagingDirectory
    }
    throw
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $designer
    Release-ComObject $component
    Release-ComObject $book
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-OwnedExcelProcessIfNeeded -ProcessId $excelProcessId
}

Copy-Item -LiteralPath $stagedFrm -Destination $frmPath -Force
Copy-Item -LiteralPath $stagedFrx -Destination $frxPath -Force
$rows | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding utf8

Write-ExportLog INFO 'Owner-edited V2 geometry exported and promoted to tracked artifacts.' @{
    controlsAndPages = $rows.Count
    geometryChanges = $geometryChanges
    frm = $frmPath
    frx = $frxPath
    manifest = $manifestPath
    backupDirectory = $backupDirectory
}
