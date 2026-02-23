# ==============================================================================
# Скрипт автоматического внедрения customUI14.xml в файл Excel (.xlsm)
# Заменяет необходимость ручного использования Office RibbonX Editor
# ==============================================================================
param (
    [string]$ExcelFilePath = ".\CreateOrder.xlsm",
    [string]$XmlFilePath = ".\resources\customUI14.xml"
)

Write-Host "Начинаем интеграцию Ribbon XML..." -ForegroundColor Cyan

# Проверки файлов
if (-Not (Test-Path $ExcelFilePath)) { Write-Error "Файл Excel не найден: $ExcelFilePath"; exit }
if (-Not (Test-Path $XmlFilePath)) { Write-Error "Файл XML не найден: $XmlFilePath"; exit }

# Создаем временную папку
$tempDir = Join-Path $env:TEMP "ExcelRibbonPatcher_$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# 1. Переименовываем .xlsm в .zip и распаковываем
$zipPath = $ExcelFilePath -replace '\.xlsm$', '.zip'
Copy-Item $ExcelFilePath $zipPath -Force
Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force

# 2. Создаем или заменяем customUI14.xml
$customUIDir = Join-Path $tempDir "customUI"
if (-Not (Test-Path $customUIDir)) { New-Item -ItemType Directory -Path $customUIDir | Out-Null }
Copy-Item $XmlFilePath (Join-Path $customUIDir "customUI14.xml") -Force

# 3. Регистрируем связь (Relationships) в Excel, если её еще нет
$relsPath = Join-Path $tempDir "_rels\.rels"
[xml]$relsXml = Get-Content $relsPath
$ns = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
$ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")

$ribbonNode = $relsXml.SelectSingleNode("//r:Relationship[@Target='customUI/customUI14.xml']", $ns)
if ($null -eq $ribbonNode) {
    $newNode = $relsXml.CreateElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
    $newNode.SetAttribute("Id", "customUIRelID")
    $newNode.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
    $newNode.SetAttribute("Target", "customUI/customUI14.xml")
    $relsXml.DocumentElement.AppendChild($newNode) | Out-Null
    $relsXml.Save($relsPath)
}

# 4. Упаковываем обратно в ZIP и переименовываем в XLSM
Remove-Item $zipPath -Force
Compress-Archive -Path "$tempDir\*" -DestinationPath $zipPath -Force
Move-Item -Path $zipPath -Destination $ExcelFilePath -Force

# Очистка
Remove-Item $tempDir -Recurse -Force

Write-Host "Готово! Интерфейс ленты успешно обновлен в $ExcelFilePath" -ForegroundColor Green