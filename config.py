# -*- coding: utf-8 -*-
"""
Конфигурация приложения
"""

import os
from pathlib import Path

# Системные пути
APP_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = APP_DIR / "settings.json"
DATA_DIR = APP_DIR / "data"
MANUAL_LINKS_FILE = DATA_DIR / "manual_links.json"

# Создаем папки если их нет
os.makedirs(DATA_DIR, exist_ok=True)

# Конфигурация интерфейса
WINDOW_TITLE = "Умная сверка 3.0 (TicketProf ↔ Bars Tour)"
WINDOW_WIDTH = 1280
WINDOW_HEIGHT = 800

# Регулярные выражения
RE_TICKET = r'\b\d{3}-\d{10}\b'
RE_ORDER = r'\b(?:Заказ|заказ|заказа|№)\s*(\d{7,10})\b'

# Ключевые слова для классификации услуг
SERVICE_CLASSIFICATION = {
    "Hotel": ["проживание", "отель", "г-ца", "гостиница"],
    "Seat": ["выбор места", "место", "mco"],
    "Fee": ["услуга ито", "сервисный сбор", "ито", "сбор"],
    "Penalty": ["штраф", "удержания", "удержание"],
    "Flight": ["авиабилет", "билет", "возврат авиабилета", "перелет", "авиа"]
}

# Синонимы для нормализации номенклатуры
TEXT_REPLACEMENTS = {
    "г-ца": "гостиница",
    "г.": "город ",
    "улица": "ул",
    "авиабилет": "билет"
}

# Стили для Excel
EXCEL_STYLES = {
    "font_name": "Arial",
    "header_fill": "1F497D",  # Темно-синий
    "header_font_color": "FFFFFF",
    "matched_fill": "E2EFDA",  # Светло-зеленый
    "unmatched_fill": "FCE4D6",  # Светло-оранжевый
    "discrepancy_fill": "FFC7CE",  # Светло-красный
    "border_color": "D9D9D9"
}

# Карты колонок для гибридного парсинга (1-based индексы)
# Сводный файл
COL_MAP_SINGLE_TP = {"date": 1, "doc": 2, "debit": 3, "credit": 4}
COL_MAP_SINGLE_BT = {"date": 8, "doc": 9, "amount": 10, "profit": 12, "net": 13}

# Раздельные файлы
COL_MAP_DOUBLE_TP = {"date": 1, "doc": 2, "debit": 3, "credit": 4}
COL_MAP_DOUBLE_BT = {"date": 1, "doc": 2, "amount": 3, "profit": 5, "net": 6}

