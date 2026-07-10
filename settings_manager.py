# -*- coding: utf-8 -*-
"""
Модуль управления настройками приложения
"""

import json
import os
from pathlib import Path
from datetime import datetime
from config import SETTINGS_FILE

def load_settings() -> str:
    """
    Загружает настройки из settings.json.
    Возвращает путь к папке по умолчанию (по умолчанию папка Downloads пользователя).
    """
    default_dir = str(Path.home() / "Downloads")
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                return settings.get('default_folder', default_dir)
    except Exception:
        pass
    return default_dir

def save_settings(default_folder: str) -> None:
    """
    Сохраняет настройки в settings.json
    """
    try:
        settings = {
            'default_folder': default_folder,
            'last_update': datetime.now().isoformat()
        }
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
