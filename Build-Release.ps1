<#
.SYNOPSIS
    Сборщик релизной версии проекта CreateOrder.
.DESCRIPTION
    Применяет двойную бинарную защиту: Unviewable Project (DPx) и Ghost Module (скрытие ключевых модулей).
    Генерирует готовый файл с меткой даты и времени (например, CreateOrder_Release_20260225_153000.xlsm).
#>

param(
    [string]$SourceFile = "CreateOrder.xlsm"
)

# Переключаем кодировку консоли на UTF-8 для корректного вывода русских букв
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Генерируем динамическое имя выходного файла с датой и временем
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "CreateOrder_Release_$timestamp.xlsm"

Write-Host "=== Запуск сборки защищенного релиза ===" -ForegroundColor Cyan

# Проверка наличия исходного файла
if (-not (Test-Path $SourceFile)) {
    Write-Host "[X] Ошибка: Исходный файл $SourceFile не найден!" -ForegroundColor Red
    exit
}

# 1. Подготовка временных директорий
$tempDir = Join-Path $env:TEMP "CreateOrderBuild_$timestamp"
$tempZip = Join-Path $tempDir "temp_archive.zip"
$extractDir = Join-Path $tempDir "extracted"

New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
Copy-Item -Path $SourceFile -Destination $tempZip -Force

Write-Host "[1/4] Распаковка архива xlsm..." -ForegroundColor Yellow
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $extractDir)

# 2. Поиск vbaProject.bin
$vbaBinPath = Join-Path $extractDir "xl\vbaProject.bin"
if (-not (Test-Path $vbaBinPath)) {
    Write-Host "[X] Ошибка: Файл vbaProject.bin не найден в архиве!" -ForegroundColor Red
    Remove-Item $tempDir -Recurse -Force
    exit
}

# 3. Бинарный патчинг (Raw Byte Patching)
Write-Host "[2/4] Применение бинарной защиты..." -ForegroundColor Yellow

# Читаем как текст в кодировке Default (ANSI/Windows-1251), чтобы сохранить 1-байтовые символы
$bytes = [System.IO.File]::ReadAllBytes($vbaBinPath)
$encoding = [System.Text.Encoding]::GetEncoding(1252) 
$text = $encoding.GetString($bytes)

# --- Слой 1: Unviewable Project (DPB -> DPx) ---
if ($text -match "DPB=") {
    $text = $text -replace "DPB=", "DPx="
    Write-Host "  -> [Unviewable]: Успешно (DPB -> DPx)" -ForegroundColor Green
} else {
    Write-Host "  -> [Unviewable]: Сигнатура DPB= не найдена (возможно, уже защищен)" -ForegroundColor DarkYellow
}

# --- Слой 2: Ghost Modules (Скрытие из дерева) ---
# ПОЛНЫЙ СПИСОК МОДУЛЕЙ ДЛЯ СКРЫТИЯ (Защита от запуска через Alt+F8)
$modulesToHide = @(
    "modActivation",             # Логика лицензии
    "mdlRibbonHandlers",         # Вызовы проверок с ленты
    "mdlMainExport",             # Основной приказ
    "mdlRaportExport",           # Рапорты
    "mdlSpravkaExport",          # Справки ДСО
    "mdlRiskExport",             # Приказ за риск
    "mdlUniversalPaymentExport", # Надбавки
    "mdlFRPExport",              # Отчеты Алушта
    "mdlWordImport",             # Импорт рапортов
    "MdlBackup",                 # Бэкапер
	"frmAbout"
)

foreach ($modName in $modulesToHide) {
    $searchStr = "Module=$modName"
    
    if ($text.Contains($searchStr)) {
        # Создаем строку из пробелов той же длины
        $spaces = " " * $searchStr.Length
        $text = $text.Replace($searchStr, $spaces)
        Write-Host "  -> [Ghosting]: Модуль '$modName' скрыт." -ForegroundColor Green
    } else {
        Write-Host "  -> [Ghosting]: Модуль '$modName' не найден в потоке PROJECT." -ForegroundColor DarkYellow
    }
}

# Сохраняем патч обратно в байты
$newBytes = $encoding.GetBytes($text)
[System.IO.File]::WriteAllBytes($vbaBinPath, $newBytes)

# 4. Обратная запаковка
Write-Host "[3/4] Запаковка релизного архива..." -ForegroundColor Yellow
Remove-Item $tempZip -Force
[System.IO.Compression.ZipFile]::CreateFromDirectory($extractDir, $tempZip)

# 5. Перенос результата
Write-Host "[4/4] Финализация..." -ForegroundColor Yellow
Copy-Item -Path $tempZip -Destination $OutputFile -Force

# Очистка мусора
Remove-Item $tempDir -Recurse -Force

Write-Host "=== ГОТОВО! ===" -ForegroundColor Cyan
Write-Host "Защищенный файл создан: $OutputFile" -ForegroundColor Green