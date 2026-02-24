@echo off
chcp 65001 > nul
echo Запуск системы сборки...
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0Build-Release.ps1"
pause