@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Установка Pillow (PIL) напрямую...

if not exist .venv\Scripts\python.exe (
    echo Виртуальное окружение не найдено! Запустите setup.bat.
    pause
    exit /b 1
)

echo Запуск: .venv\Scripts\python.exe -m pip install Pillow
.venv\Scripts\python.exe -m pip install Pillow

if %errorlevel% neq 0 (
    echo.
    echo ❌ Произошла ошибка при установке Pillow через pip!
    echo Проверьте подключение к интернету.
) else (
    echo.
    echo ✅ Установка завершена. Проверка импорта PIL:
    .venv\Scripts\python.exe -c "import PIL; print('Успешно! Версия PIL:', PIL.__version__)"
)
pause
