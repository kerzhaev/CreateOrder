[CmdletBinding()]
param(
    [string]$WorkbookPath = ".\CreateOrder.xlsm",
    [string]$ModulePath = ".\CreateOrder.xlsm.modules\mdlZP12Validation.bas",
    [string]$RibbonHandlersPath = ".\CreateOrder.xlsm.modules\mdlRibbonHandlers.bas",
    [string]$SourceTemplatePath = "",
    [string]$RunTemplatePath = ".\ZP12_TEST_RUN.xlsx",
    [string]$Pass1SnapshotPath = ".\ZP12_TEST_RUN_PASS1.xlsx"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Resolve-ProjectPath {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot $PathValue))
}

function Invoke-Python {
    param([string]$ScriptText)

    $tempFile = Join-Path $env:TEMP ("zpt12_" + [System.Guid]::NewGuid().ToString("N") + ".py")
    try {
        Set-Content -Path $tempFile -Value $ScriptText -Encoding UTF8
        $output = & python $tempFile 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            throw ($output -join [Environment]::NewLine)
        }
        return ($output -join [Environment]::NewLine)
    }
    finally {
        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    }
}

function Get-DefaultSourceTemplate {
    $candidate = Get-ChildItem -Path $PSScriptRoot -File | Where-Object {
        $_.Name -like "*_TEST.xlsx" -and
        $_.Name -notlike "*RUN*" -and
        $_.Name -notlike "*PASS1*"
    } | Sort-Object Name | Select-Object -First 1

    if ($null -eq $candidate) {
        throw "Не найден исходный тестовый шаблон (*_TEST.xlsx)."
    }

    return $candidate.FullName
}

function Invoke-ZP12MacroRun {
    param(
        [string]$WorkbookFullPath,
        [string]$ModuleFullPath,
        [string]$RibbonHandlersFullPath,
        [string]$TemplateFullPath
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    try {
        $macroWb = $excel.Workbooks.Open($WorkbookFullPath)
        $vbProject = $macroWb.VBProject

        try {
            $component = $vbProject.VBComponents.Item("mdlZP12Validation")
            $vbProject.VBComponents.Remove($component)
        }
        catch {
        }

        try {
            $component = $vbProject.VBComponents.Item("mdlRibbonHandlers")
            $vbProject.VBComponents.Remove($component)
        }
        catch {
        }

        $null = $vbProject.VBComponents.Import($ModuleFullPath)
        $null = $vbProject.VBComponents.Import($RibbonHandlersFullPath)
        $macroWb.Save()

        $excel.Run("CreateOrder.xlsm!mdlZP12Validation.ValidateZP12Template", $TemplateFullPath, $true)

        $templateWb = $null
        foreach ($wb in $excel.Workbooks) {
            if ($wb.FullName -eq $TemplateFullPath) {
                $templateWb = $wb
                break
            }
        }

        if ($null -ne $templateWb) {
            $templateWb.Save()
            $templateWb.Close($false)
        }

        $macroWb.Close($true)
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-ZP12State {
    param([string]$WorkbookFullPath)

    $python = @"
import json
from pathlib import Path
from openpyxl import load_workbook

path = Path(r'''$WorkbookFullPath''')
wb = load_workbook(path)
history_name = 'История_проверок_ZP12'
if history_name not in wb.sheetnames:
    raise SystemExit('History sheet not found')

main_ws = wb[wb.sheetnames[0]]
history_ws = wb[history_name]

highlights = []
for row in range(4, 40):
    for col in range(2, 12):
        cell = main_ws.cell(row, col)
        fill = cell.fill
        color = None
        if fill and fill.fill_type == 'solid':
            color = fill.fgColor.rgb or fill.start_color.rgb
        if color or cell.comment:
            highlights.append({
                'cell': cell.coordinate,
                'color': color,
                'comment': cell.comment.text.strip() if cell.comment else ''
            })

headers = [history_ws.cell(1, c).value for c in range(1, history_ws.max_column + 1)]
rows = []
status_by_run = {}
for r in range(2, history_ws.max_row + 1):
    values = [history_ws.cell(r, c).value for c in range(1, history_ws.max_column + 1)]
    if any(v is not None and v != '' for v in values):
        item = dict(zip(headers, values))
        rows.append(item)
        run_id = str(item['RunID'])
        status_by_run.setdefault(run_id, {})
        status = item['Статус']
        status_by_run[run_id][status] = status_by_run[run_id].get(status, 0) + 1

print(json.dumps({
    'highlightedCells': [item['cell'] for item in highlights],
    'highlights': highlights,
    'historyRowCount': len(rows),
    'statusByRun': status_by_run
}, ensure_ascii=False))
"@

    return (Invoke-Python -ScriptText $python | ConvertFrom-Json)
}

function Apply-RegressionFixes {
    param([string]$WorkbookFullPath)

    $python = @"
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

path = Path(r'''$WorkbookFullPath''')
wb = load_workbook(path)
ws = wb[wb.sheetnames[0]]

ws['B5'] = 21463
ws['G7'] = 'Владимирович'
ws['I11'] = datetime(2023, 10, 21)
ws['J11'] = datetime(2023, 10, 25)
ws['I11'].number_format = 'dd.mm.yyyy'
ws['J11'].number_format = 'dd.mm.yyyy'

wb.save(path)
print('ok')
"@

    Invoke-Python -ScriptText $python | Out-Null
}

function Get-StatusCount {
    param(
        [object]$State,
        [string]$RunId,
        [string]$StatusName
    )

    $run = $State.statusByRun.PSObject.Properties[$RunId]
    if ($null -eq $run) {
        return 0
    }

    $status = $run.Value.PSObject.Properties[$StatusName]
    if ($null -eq $status) {
        return 0
    }

    return [int]$status.Value
}

function Assert-Equal {
    param(
        [string]$Label,
        [object]$Expected,
        [object]$Actual
    )

    if ($Expected -ne $Actual) {
        throw "${Label}: ожидалось '$Expected', получено '$Actual'."
    }
}

function Assert-SetEqual {
    param(
        [string]$Label,
        [string[]]$Expected,
        [object[]]$Actual
    )

    $expectedSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$Expected)
    $actualSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$Actual)

    if (-not $expectedSet.SetEquals($actualSet)) {
        $expectedText = ($Expected | Sort-Object) -join ", "
        $actualText = ([string[]]$Actual | Sort-Object) -join ", "
        throw "${Label}: ожидалось [$expectedText], получено [$actualText]."
    }
}

$workbookFullPath = Resolve-ProjectPath $WorkbookPath
$moduleFullPath = Resolve-ProjectPath $ModulePath
$ribbonHandlersFullPath = Resolve-ProjectPath $RibbonHandlersPath
$sourceTemplateFullPath = if ([string]::IsNullOrWhiteSpace($SourceTemplatePath)) { Get-DefaultSourceTemplate } else { Resolve-ProjectPath $SourceTemplatePath }
$runTemplateFullPath = Resolve-ProjectPath $RunTemplatePath
$pass1SnapshotFullPath = Resolve-ProjectPath $Pass1SnapshotPath

if (-not (Test-Path $workbookFullPath)) { throw "Не найден workbook: $workbookFullPath" }
if (-not (Test-Path $moduleFullPath)) { throw "Не найден модуль: $moduleFullPath" }
if (-not (Test-Path $ribbonHandlersFullPath)) { throw "Не найден ribbon-модуль: $ribbonHandlersFullPath" }
if (-not (Test-Path $sourceTemplateFullPath)) { throw "Не найден шаблон: $sourceTemplateFullPath" }

Write-Host "=== ZP12 regression ===" -ForegroundColor Cyan
Write-Host "Workbook: $workbookFullPath"
Write-Host "Module:   $moduleFullPath"
Write-Host "Ribbon:   $ribbonHandlersFullPath"
Write-Host "Fixture:  $sourceTemplateFullPath"

Copy-Item -Path $sourceTemplateFullPath -Destination $runTemplateFullPath -Force
if (Test-Path $pass1SnapshotFullPath) {
    Remove-Item -Path $pass1SnapshotFullPath -Force
}

Write-Host "[1/5] First validation run..." -ForegroundColor Yellow
Invoke-ZP12MacroRun -WorkbookFullPath $workbookFullPath -ModuleFullPath $moduleFullPath -RibbonHandlersFullPath $ribbonHandlersFullPath -TemplateFullPath $runTemplateFullPath
$pass1State = Get-ZP12State -WorkbookFullPath $runTemplateFullPath
Copy-Item -Path $runTemplateFullPath -Destination $pass1SnapshotFullPath -Force

Write-Host "[2/5] Apply controlled fixes..." -ForegroundColor Yellow
Apply-RegressionFixes -WorkbookFullPath $runTemplateFullPath

Write-Host "[3/5] Second validation run..." -ForegroundColor Yellow
Invoke-ZP12MacroRun -WorkbookFullPath $workbookFullPath -ModuleFullPath $moduleFullPath -RibbonHandlersFullPath $ribbonHandlersFullPath -TemplateFullPath $runTemplateFullPath
$pass2State = Get-ZP12State -WorkbookFullPath $runTemplateFullPath

Write-Host "[4/5] Assert expected status transitions..." -ForegroundColor Yellow
Assert-Equal -Label "Pass1 NEW" -Expected 7 -Actual (Get-StatusCount -State $pass1State -RunId "1" -StatusName "NEW")
Assert-Equal -Label "Pass1 history rows" -Expected 7 -Actual $pass1State.historyRowCount
Assert-Equal -Label "Pass2 OPEN" -Expected 4 -Actual (Get-StatusCount -State $pass2State -RunId "2" -StatusName "OPEN")
Assert-Equal -Label "Pass2 RESOLVED" -Expected 3 -Actual (Get-StatusCount -State $pass2State -RunId "2" -StatusName "RESOLVED")
Assert-Equal -Label "Pass2 history rows" -Expected 14 -Actual $pass2State.historyRowCount

$expectedPass2Highlights = @("C6", "I8", "J8", "I9", "J9", "B10", "C10")
Assert-SetEqual -Label "Pass2 highlighted cells" -Expected $expectedPass2Highlights -Actual $pass2State.highlightedCells

Write-Host "[5/5] Summary" -ForegroundColor Yellow
Write-Host "Pass1: NEW=$(Get-StatusCount -State $pass1State -RunId '1' -StatusName 'NEW'), rows=$($pass1State.historyRowCount)" -ForegroundColor Green
Write-Host "Pass2: OPEN=$(Get-StatusCount -State $pass2State -RunId '2' -StatusName 'OPEN'), RESOLVED=$(Get-StatusCount -State $pass2State -RunId '2' -StatusName 'RESOLVED'), rows=$($pass2State.historyRowCount)" -ForegroundColor Green
Write-Host "Remaining highlighted cells: $(([string[]]$pass2State.highlightedCells | Sort-Object) -join ', ')" -ForegroundColor Green
Write-Host "Run file: $runTemplateFullPath" -ForegroundColor Green
Write-Host "Pass1 snapshot: $pass1SnapshotFullPath" -ForegroundColor Green
Write-Host "RESULT: PASS" -ForegroundColor Cyan
