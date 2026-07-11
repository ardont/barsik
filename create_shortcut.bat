@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Создание ярлыка на Рабочем столе...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0create_shortcut.ps1"
pause
