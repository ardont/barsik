@echo off
chcp 65001 > nul
title Умная сверка 3.0 - Запуск

echo ====================================================
echo   Запуск приложения «Умная сверка 3.0»
echo ====================================================
echo.

:: 1. Проверка наличия Python
python --version >nul 2>nul
if %errorlevel% equ 0 goto python_ok

echo [ОШИБКА] Python не найден в системе!
echo Пожалуйста, скачайте и установите Python 3.10 или 3.11 с официального сайта:
echo https://www.python.org/downloads/
echo При установке ОБЯЗАТЕЛЬНО отметьте галочку "Add Python to PATH" (Добавить Python в PATH).
echo.
pause
exit /b 1

:python_ok

:: 2. Создание виртуального окружения, если его нет
if exist .venv goto venv_ok
echo [ИНФО] Виртуальное окружение .venv не найдено. Создание...
python -m venv .venv
if %errorlevel% equ 0 goto venv_ok
echo [ОШИБКА] Не удалось создать виртуальное окружение .venv!
pause
exit /b 1

:venv_ok

:: 3. Проверка установленных зависимостей
echo [ИНФО] Проверка целостности библиотек...
.venv\Scripts\python.exe -c "import customtkinter, pandas, openpyxl, docx, PIL" >nul 2>nul
if %errorlevel% equ 0 goto run_app

echo [ИНФО] Установка или восстановление необходимых библиотек (это может занять 1-2 минуты)...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if %errorlevel% equ 0 goto run_app
echo [ОШИБКА] Не удалось установить библиотеки! Проверьте подключение к интернету.
pause
exit /b 1

:run_app
:: 4. Запуск программы
echo [ИНФО] Запуск «Умной сверки»...
.venv\Scripts\python.exe main.py
echo.
echo [ИНФО] Программа завершила работу с кодом %errorlevel%.
pause
