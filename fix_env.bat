@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Сброс и восстановление окружения...
if exist .venv (
    echo Удаление старой папки .venv...
    rd /s /q .venv
)
call run.bat
