param(
    [string]$WorkbookPath = ".\CreateOrder.xlsm",
    [string]$HelperPath = ".\CreateOrder.xlsm.modules\mdlHelper.bas"
)

$ErrorActionPreference = 'Stop'

$WorkbookFullPath = (Resolve-Path $WorkbookPath).Path
$HelperFullPath = (Resolve-Path $HelperPath).Path
$TempHelperPath = Join-Path $env:TEMP "mdlHelper_fio_test_import.bas"
$TempWorkbookPath = Join-Path $env:TEMP ("CreateOrder_FIO_Test_{0}.xlsm" -f ([guid]::NewGuid().ToString("N")))

function Stop-ExcelProcesses {
    Get-Process Excel -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 500
}

$cases = @(
    @{
        Source = "Иванов Иван Иванович"
        Dative = "Иванову Ивану Ивановичу"
        InitialsDative = "И.И. Иванову"
        InitialsNominative = "И.И. Иванов"
    },
    @{
        Source = "Петрова Анна Сергеевна"
        Dative = "Петровой Анне Сергеевне"
        InitialsDative = "А.С. Петровой"
        InitialsNominative = "А.С. Петрова"
    },
    @{
        Source = "Ильин Илья Сергеевич"
        Dative = "Ильину Илье Сергеевичу"
        InitialsDative = "И.С. Ильину"
        InitialsNominative = "И.С. Ильин"
    },
    @{
        Source = "Любимова Любовь Ивановна"
        Dative = "Любимовой Любови Ивановне"
        InitialsDative = "Л.И. Любимовой"
        InitialsNominative = "Л.И. Любимова"
    },
    @{
        Source = "Кравец Петр Иванович"
        Dative = "Кравцу Петру Ивановичу"
        InitialsDative = "П.И. Кравцу"
        InitialsNominative = "П.И. Кравец"
    },
    @{
        Source = "Сидоренко Петр Иванович"
        Dative = "Сидоренко Петру Ивановичу"
        InitialsDative = "П.И. Сидоренко"
        InitialsNominative = "П.И. Сидоренко"
    },
    @{
        Source = "Белый Никита Сергеевич"
        Dative = "Белому Никите Сергеевичу"
        InitialsDative = "Н.С. Белому"
        InitialsNominative = "Н.С. Белый"
    },
    @{
        Source = "Соколов Лев Павлович"
        Dative = "Соколову Льву Павловичу"
        InitialsDative = "Л.П. Соколову"
        InitialsNominative = "Л.П. Соколов"
    },
    @{
        Source = "Павлов Павел Ильич"
        Dative = "Павлову Павлу Ильичу"
        InitialsDative = "П.И. Павлову"
        InitialsNominative = "П.И. Павлов"
    }
)

Stop-ExcelProcesses

$helperContent = [System.IO.File]::ReadAllText($HelperFullPath, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($TempHelperPath, $helperContent, [System.Text.Encoding]::GetEncoding(1251))
Copy-Item -LiteralPath $WorkbookFullPath -Destination $TempWorkbookPath -Force

$excel = $null
$workbook = $null
$failures = New-Object System.Collections.Generic.List[string]

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($TempWorkbookPath, 0, $false)
    $vbProject = $workbook.VBProject

    try {
        $component = $vbProject.VBComponents.Item("mdlHelper")
        $vbProject.VBComponents.Remove($component)
    } catch {}

    $null = $vbProject.VBComponents.Import($TempHelperPath)
    $macroPrefix = "'$($workbook.Name)'!"

    for ($i = 0; $i -lt $cases.Count; $i++) {
        $actualDative = [string]$excel.Run($macroPrefix + "SklonitFIO", $cases[$i].Source)
        $actualInitialsDative = [string]$excel.Run($macroPrefix + "GetFIOWithInitials", $cases[$i].Source)
        $actualInitialsNominative = [string]$excel.Run($macroPrefix + "GetFIOWithInitialsImenitelny", $cases[$i].Source)

        Write-Host "CASE: $($cases[$i].Source)"
        Write-Host "  Dative:            $actualDative"
        Write-Host "  Initials Dative:   $actualInitialsDative"
        Write-Host "  Initials Nominative: $actualInitialsNominative"

        if ($actualDative -ne $cases[$i].Dative) {
            $failures.Add("Dative mismatch for '$($cases[$i].Source)': expected '$($cases[$i].Dative)', got '$actualDative'")
        }
        if ($actualInitialsDative -ne $cases[$i].InitialsDative) {
            $failures.Add("Initials dative mismatch for '$($cases[$i].Source)': expected '$($cases[$i].InitialsDative)', got '$actualInitialsDative'")
        }
        if ($actualInitialsNominative -ne $cases[$i].InitialsNominative) {
            $failures.Add("Initials nominative mismatch for '$($cases[$i].Source)': expected '$($cases[$i].InitialsNominative)', got '$actualInitialsNominative'")
        }
    }

    $workbook.Close($false)
    $workbook = $null

    if ($failures.Count -gt 0) {
        $failures | ForEach-Object { Write-Host $_ -ForegroundColor Red }
        throw "FIO declension test failed."
    }

    Write-Host "RESULT: PASS" -ForegroundColor Green
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if (Test-Path $TempHelperPath) {
        Remove-Item -LiteralPath $TempHelperPath -Force
    }
    if (Test-Path $TempWorkbookPath) {
        Remove-Item -LiteralPath $TempWorkbookPath -Force
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-ExcelProcesses
}
