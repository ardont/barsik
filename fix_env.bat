@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Удаление старого окружения и переустановка...
if exist .venv rd /s /q .venv
call run.bat
