@echo off
chcp 65001 > nul
title Умная сверка 3.0 - Запуск

echo ====================================================
echo   Запуск приложения «Умная сверка 3.0»
echo ====================================================
echo.

:: 1. Проверка наличия Python
python --version >nul 2>nul
if %errorlevel% neq 0 (
    echo [ОШИБКА] Python не найден в системе!
    echo Пожалуйста, скачайте и установите Python 3.10 или 3.11 с официального сайта:
    echo https://www.python.org/downloads/
    echo При установке ОБЯЗАТЕЛЬНО отметьте галочку "Add Python to PATH" (Добавить Python в PATH).
    echo.
    pause
    exit /b 1
)

:: 2. Создание виртуального окружения, если его нет
if not exist .venv (
    echo [ИНФО] Виртуальное окружение .venv не найдено. Создание...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ОШИБКА] Не удалось создать виртуальное окружение .venv!
        pause
        exit /b 1
    )
)

:: 3. Проверка установленных зависимостей
echo [ИНФО] Проверка целостности библиотек...
.venv\Scripts\python.exe -c "import customtkinter, pandas, openpyxl, docx, PIL" >nul 2>nul
if %errorlevel% neq 0 (
    echo [ИНФО] Установка или восстановление необходимых библиотек (это может занять 1-2 минуты)...
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [ОШИБКА] Не удалось установить библиотеки! Проверьте подключение к интернету.
        pause
        exit /b 1
    )
)

:: 4. Запуск программы
echo [ИНФО] Запуск «Умной сверки»...
.venv\Scripts\python.exe main.py
if %errorlevel% neq 0 (
    echo.
    echo [ОШИБКА] Программа завершилась с ошибкой.
    echo.
    pause
)
