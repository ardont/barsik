@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Запуск диагностики окружения...

if not exist .venv\Scripts\python.exe (
    echo Виртуальное окружение не найдено. Запуск настройки...
    call run.bat
)

if exist .venv\Scripts\python.exe (
    echo Запуск диагностики через .venv...
    .venv\Scripts\python.exe check_setup.py
) else (
    echo [ОШИБКА] Не удалось запустить диагностику.
)
pause
