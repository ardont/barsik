@echo off
chcp 65001 > nul
if not exist .venv (
    echo [Умная сверка 3.0] Виртуальное окружение не обнаружено! Запустите setup.bat.
    pause
    exit /b 1
)
echo [Умная сверка 3.0] Обновление библиотек...
.venv\Scripts\python.exe -m pip install -r requirements.txt --upgrade
if %errorlevel% neq 0 (
    echo Ошибка при обновлении библиотек!
    pause
    exit /b %errorlevel%
)
echo Обновление завершено!
pause
