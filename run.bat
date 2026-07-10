@echo off
chcp 65001 > nul
if not exist .venv (
    echo [Умная сверка 3.0] Виртуальное окружение не обнаружено!
    echo Запускаем установку через setup.bat...
    call setup.bat
)
.venv\Scripts\python.exe main.py
if %errorlevel% neq 0 (
    echo Ошибка во время выполнения программы.
    pause
)
