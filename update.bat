@echo off
chcp 65001 > nul
if not exist .venv (
    echo [Умная сверка 3.0] Виртуальное окружение не обнаружено! Запустите setup.bat.
    pause
    exit /b 1
)
call .venv\Scripts\activate
echo [Умная сверка 3.0] Обновление библиотек...
pip install -r requirements.txt --upgrade
echo Обновление завершено!
pause
