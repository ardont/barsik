# -*- coding: utf-8 -*-
"""
Диагностический скрипт для проверки виртуального окружения
"""

import sys
import os

log = []
log.append("=== ДИАГНОСТИКА ОКРУЖЕНИЯ БАРСИК ===")
log.append(f"Текущая папка: {os.getcwd()}")
log.append(f"Текущий запускаемый Python: {sys.executable}")
log.append(f"Версия Python: {sys.version}")

# Проверяем наличие Pillow
try:
    import PIL
    from PIL import Image
    log.append(f"✅ Pillow (PIL) успешно импортирован! Версия: {PIL.__version__}")
except ImportError as e:
    log.append(f"❌ Ошибка импорта Pillow (PIL): {e}")

# Проверяем наличие других библиотек
for lib in ["customtkinter", "pandas", "openpyxl", "docx"]:
    try:
        __import__(lib)
        log.append(f"✅ Библиотека '{lib}' успешно импортирована.")
    except ImportError as e:
        log.append(f"❌ Ошибка импорта '{lib}': {e}")

# Проверяем содержимое папки .venv
venv_dir = os.path.join(os.getcwd(), ".venv")
if os.path.exists(venv_dir):
    log.append(f"✅ Папка .venv существует по пути: {venv_dir}")
    python_exe = os.path.join(venv_dir, "Scripts", "python.exe")
    if os.path.exists(python_exe):
        log.append(f"  ✅ Файл python.exe в .venv существует.")
    else:
        log.append(f"  ❌ Файл python.exe в .venv отсутствует!")
else:
    log.append("❌ Папка .venv не найдена в текущей папке!")

with open("debug_log.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(log))

print("\n".join(log))
print("\nДиагностический лог сохранен в файл debug_log.txt")
