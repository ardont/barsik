@echo off
chcp 65001 > nul
echo === ДИАГНОСТИКА СРЕДЫ ЗАПУСКА ===
python --version
echo Виртуальное окружение (.venv):
if exist .venv (echo    - Найдено) else (echo    - НЕ НАЙДЕНО)
echo.
pause
