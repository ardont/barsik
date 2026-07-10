@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Запуск диагностики окружения...

if exist .venv\Scripts\python.exe (
    echo Запуск через виртуальное окружение .venv...
    .venv\Scripts\python.exe check_setup.py
) else (
    echo Виртуальное окружение не найдено. Запуск через глобальный Python...
    python check_setup.py
)

if %errorlevel% neq 0 (
    echo Критическая ошибка при работе скрипта диагностики!
)
pause
