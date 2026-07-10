# -*- coding: utf-8 -*-
"""
Модуль очистки и нормализации текстовых данных
"""

import re
from typing import Set
from config import SERVICE_CLASSIFICATION, TEXT_REPLACEMENTS, RE_TICKET, RE_ORDER

def classify_service(text: str) -> str:
    """
    Классифицирует тип услуги на основе ключевых слов в описании
    """
    if not text:
        return "Other"
    
    t = text.lower()
    for category, keywords in SERVICE_CLASSIFICATION.items():
        if any(kw in t for kw in keywords):
            return category
            
    return "Other"

def clean_text(text: str) -> str:
    """
    Очищает описание услуги от мусора для сравнения:
    - Приводит к нижнему регистру
    - Применяет замены сокращений (г-ца -> гостиница)
    - Вырезает даты и периоды времени
    - Вырезает имена пассажиров и английские слова в скобках
    - Удаляет номера билетов и заказов
    - Удаляет знаки препинания и лишние пробелы
    """
    if not text:
        return ""
        
    t = text.lower()
    
    # Применяем замены
    for orig, rep in TEXT_REPLACEMENTS.items():
        t = t.replace(orig, rep)
        
    # Вырезаем даты формата DD.MM.YYYY или DD.MM.YY
    t = re.sub(r'\b\d{2}\.\d{2}\.(?:\d{4}|\d{2})\b', '', t)
    
    # Вырезаем диапазоны дат типа "с DD.MM по DD.MM"
    t = re.sub(r'\bс\s+\d{2}\.\d{2}\s+по\s+\d{2}\.\d{2}\b', '', t)
    t = re.sub(r'\bс\s+\d{2}\.\d{2}\.\d{2,4}\s+по\s+\d{2}\.\d{2}\.\d{2,4}\b', '', t)
    
    # Вырезаем инфо о пассажирах на латинице в скобках, например: (SOLOVEVA/MISHEL MS...)
    t = re.sub(r'\([a-z\s\/]+:?\s*(?:№)?\s*\d*(?:-\d*)?\)', '', t)
    t = re.sub(r'\([a-z0-9\s\/\-\:\.\,]+\)', '', t)
    
    # Вырезаем шаблоны имен типа SOLOVEVA/MISHEL
    t = re.sub(r'\b[a-z]+/[a-z]+\b', '', t)
    
    # Вырезаем номера билетов (13 цифр через дефис)
    t = re.sub(r'\b\d{3}-\d{10}\b', '', t)
    
    # Вырезаем номера заказов и прочие числовые ID
    t = re.sub(r'\b(?:заказ|заказа|№)?\s*\d+\b', '', t)
    
    # Удаляем пунктуацию и оставляем только буквы и пробелы
    t = re.sub(r'[^\w\s]', ' ', t)
    
    # Схлопываем лишние пробелы
    t = re.sub(r'\s+', ' ', t).strip()
    
    return t

def extract_identifiers(text: str) -> Set[str]:
    """
    Извлекает множество уникальных ID (билетов, MCO, заказов) из описания
    """
    if not text:
        return set()
        
    # Ищем номера билетов
    tickets = re.findall(RE_TICKET, text)
    
    # Ищем номера заказов (числа от 7 до 10 цифр)
    orders = re.findall(RE_ORDER, text)
    
    return set(tickets + orders)
