@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Полный сброс и пересоздание окружения...

:: Удаляем старое окружение
if exist .venv (
    echo Удаление старой папки .venv...
    rd /s /q .venv
)

echo Создание нового виртуального окружения...
python -m venv .venv
if %errorlevel% neq 0 (
    echo Ошибка при создании виртуального окружения! Проверьте, установлен ли Python.
    pause
    exit /b 1
)

echo Установка библиотек (это займет около 1-2 минут)...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo Произошла ошибка при установке библиотек.
    pause
    exit /b 1
)

echo.
echo Список установленных библиотек в окружении:
.venv\Scripts\pip.exe list

echo.
echo === Сброс завершен! Запустите run.bat или create_shortcut.bat ===
pause
