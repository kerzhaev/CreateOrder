<#
.SYNOPSIS
    ������� �������� ������ ������� CreateOrder.
.DESCRIPTION
    Подготавливает защищённую релизную сборку: штатный пароль на VBA-проект и Ghost Module (скрытие важных модулей).
    ���������� ������� ���� � ������ ���� � ������� (��������, CreateOrder_Release_20260225_153000.xlsm).
#>

param(
    [string]$SourceFile = "CreateOrder.xlsm"
)

# ����������� ��������� ������� �� UTF-8 ��� ����������� ������ ������� ����
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ���������� ������������ ��� ��������� ����� � ����� � ��������
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$releaseDir = Join-Path (Get-Location) "CreateOrderReleases"
$OutputFile = Join-Path $releaseDir "CreateOrder_Release_$timestamp.xlsm"

Write-Host "=== ������ ������ ����������� ������ ===" -ForegroundColor Cyan

# ������� ���������� ����� ��� �������� �������� �������
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

# �������� ������� ��������� �����
if (-not (Test-Path $SourceFile)) {
    Write-Host "[X] ������: �������� ���� $SourceFile �� ������!" -ForegroundColor Red
    exit
}

# 1. ���������� ��������� ����������
$tempDir = Join-Path $env:TEMP "CreateOrderBuild_$timestamp"
$tempZip = Join-Path $tempDir "temp_archive.zip"
$extractDir = Join-Path $tempDir "extracted"

New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
Copy-Item -Path $SourceFile -Destination $tempZip -Force

Write-Host "[1/4] ���������� ������ xlsm..." -ForegroundColor Yellow
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $extractDir)

# 2. ����� vbaProject.bin
$vbaBinPath = Join-Path $extractDir "xl\vbaProject.bin"
if (-not (Test-Path $vbaBinPath)) {
    Write-Host "[X] ������: ���� vbaProject.bin �� ������ � ������!" -ForegroundColor Red
    Remove-Item $tempDir -Recurse -Force
    exit
}

# 3. �������� ������� (Binary Byte Patching)
Write-Host "[2/4] ���������� �������� ������..." -ForegroundColor Yellow

$bytes = [System.IO.File]::ReadAllBytes($vbaBinPath)
$ascii = [System.Text.Encoding]::ASCII

function Replace-ByteSequence {
    param(
        [byte[]]$Buffer,
        [byte[]]$Search,
        [byte[]]$Replace
    )

    if ($Search.Length -ne $Replace.Length) {
        throw "Search and replace sequences must have the same length."
    }

    $matches = 0
    for ($i = 0; $i -le $Buffer.Length - $Search.Length; $i++) {
        $found = $true
        for ($j = 0; $j -lt $Search.Length; $j++) {
            if ($Buffer[$i + $j] -ne $Search[$j]) {
                $found = $false
                break
            }
        }

        if ($found) {
            for ($j = 0; $j -lt $Replace.Length; $j++) {
                $Buffer[$i + $j] = $Replace[$j]
            }
            $matches++
        }
    }

    return $matches
}

# --- Слой 1: штатный пароль на VBA-проект, без DPx-хаков ---

# --- ���� 2: Ghost Modules (������� �� ������) ---
# ������ ������ ������� ��� ������� (������ �� ������� ����� Alt+F8)
$modulesToHide = @(
    "modActivation",             # ������ ��������
    "mdlRibbonHandlers",         # ������ �������� � �����
    "mdlMainExport",             # �������� ������
    "mdlRaportExport",           # �������
    "mdlSpravkaExport",          # ������� ���
    "mdlRiskExport",             # ������ �� ����
    "mdlUniversalPaymentExport", # ��������
    "mdlFRPExport",              # ������ ������
    "mdlWordImport",             # ������ ��������
    "MdlBackup",                 # �������
	"frmAbout"
)

$ghostedCount = 0
foreach ($modName in $modulesToHide) {
    $searchStr = "Module=$modName"
    $searchBytes = $ascii.GetBytes($searchStr)
    $replaceBytes = $ascii.GetBytes((" " * $searchStr.Length))
    $patchedCount = Replace-ByteSequence -Buffer $bytes -Search $searchBytes -Replace $replaceBytes

    if ($patchedCount -gt 0) {
        $ghostedCount += $patchedCount
        Write-Host "  -> [Ghosting]: ������ '$modName' �����. matches=$patchedCount" -ForegroundColor Green
    } else {
        Write-Host "  -> [Ghosting]: ������ '$modName' �� ������ � ������ PROJECT." -ForegroundColor DarkYellow
    }
}

# ��������� ���� ������� � �����
[System.IO.File]::WriteAllBytes($vbaBinPath, $bytes)
Write-Host "  -> [Patch]: ghosted modules=$ghostedCount" -ForegroundColor Green

# 4. �������� ���������
Write-Host "[3/4] ��������� ��������� ������..." -ForegroundColor Yellow
Remove-Item $tempZip -Force
[System.IO.Compression.ZipFile]::CreateFromDirectory($extractDir, $tempZip)

# 5. ������� ����������
Write-Host "[4/4] �����������..." -ForegroundColor Yellow
Copy-Item -Path $tempZip -Destination $OutputFile -Force

# ������� ������
Remove-Item $tempDir -Recurse -Force

Write-Host "=== ������! ===" -ForegroundColor Cyan
Write-Host "���������� ���� ������: $OutputFile" -ForegroundColor Green
