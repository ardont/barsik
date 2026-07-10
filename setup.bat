@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Создание виртуального окружения...
python -m venv .venv
if %errorlevel% neq 0 (
    echo Ошибка при создании виртуального окружения! Убедитесь, что Python установлен.
    pause
    exit /b %errorlevel%
)

echo [Умная сверка 3.0] Обновление pip и установка зависимостей...
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Ошибка при установке библиотек! Проверьте подключение к сети.
    pause
    exit /b %errorlevel%
)

echo [Умная сверка 3.0] Настройка завершена! Для запуска используйте run.bat
pause
