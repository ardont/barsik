@echo off
chcp 65001 > nul
echo [Умная сверка 3.0] Переустановка библиотеки изображений (Pillow)...
.venv\Scripts\python.exe -m pip install --force-reinstall pillow
pause
