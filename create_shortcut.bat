@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Создание ярлыка программы...

:: Проверяем наличие Python и Pillow для генерации иконки
.venv\Scripts\python.exe -c "import PIL" >nul 2>nul
if %errorlevel% neq 0 (
    echo [ИНФО] Настройка окружения для генерации иконки...
    if not exist .venv (
        python -m venv .venv
    )
    .venv\Scripts\python.exe -m pip install Pillow
)

:: Конвертируем логотип в формат ICO
if not exist assets\barsik_logo.ico (
    echo Создаем файл иконки...
    .venv\Scripts\python.exe -c "from PIL import Image; Image.open('assets/barsik_logo.png').save('assets/barsik_logo.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])"
)

:: Запускаем PowerShell скрипт для создания ярлыка
echo Создаем ярлык на Рабочем столе...
powershell -ExecutionPolicy Bypass -File create_shortcut.ps1

if %errorlevel% neq 0 (
    echo Не удалось создать ярлык.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo Ярлык 'Барсик - Умная сверка' успешно создан на вашем Рабочем столе!
echo.
echo Чтобы закрепить программу на панели задач:
echo 1. Перетащите созданный ярлык с Рабочего стола на Панель задач Windows
echo 2. Или запустите программу через ярлык, кликните правой кнопкой мыши
echo    по значку программы на Панели задач и выберите 'Закрепить на панели задач'.
echo ================================================================
echo.
pause
