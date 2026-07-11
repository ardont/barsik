# -*- coding: utf-8 -*-
"""
Модуль управления настройками приложения
"""

import json
import os
from pathlib import Path
from datetime import datetime
from config import SETTINGS_FILE

DEFAULT_SETTINGS = {
    'default_folder': str(Path.home() / "Downloads"),
    'hotel_margin': 10.0,
    'fuzzy_threshold': 75.0,
    'enable_id_match': True,
    'enable_exact_match': True,
    'enable_fuzzy_match': True
}

def load_settings() -> dict:
    """
    Загружает настройки из settings.json.
    Возвращает словарь со всеми настройками.
    """
    settings = DEFAULT_SETTINGS.copy()
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                for k, v in DEFAULT_SETTINGS.items():
                    settings[k] = saved.get(k, v)
    except Exception:
        pass
    return settings

def save_settings(settings: dict) -> None:
    """
    Сохраняет настройки в settings.json
    """
    try:
        settings['last_update'] = datetime.now().isoformat()
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
