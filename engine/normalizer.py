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
    - Вырезает имена пассажиров и английские/русские слова в скобках
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
    
    # Вырезаем скобки с пассажирами/номерами
    t = re.sub(r'\([a-z0-9\s\/\-\:\.\,№а-я]+\)', '', t)
    
    # Вырезаем шаблоны имен типа SOLOVEVA/MISHEL
    t = re.sub(r'\b[a-z]+/[a-z]+\b', '', t)
    
    # Вырезаем номера билетов (авиа и жд)
    t = re.sub(r'\b[0-9а-яa-z]{2,4}-\d{10,14}\b', '', t)
    t = re.sub(r'\b\d{13,14}\b', '', t)
    
    # Вырезаем номера заказов и прочие числовые ID
    t = re.sub(r'\b(?:заказ|заказа|№)?\s*\d+\b', '', t)
    
    # Удаляем пунктуацию и оставляем только буквы и пробелы
    t = re.sub(r'[^\w\s]', ' ', t)
    
    # Схлопываем лишние пробелы
    t = re.sub(r'\s+', ' ', t).strip()
    
    return t

def extract_identifiers(text: str) -> Set[str]:
    """
    Извлекает множество уникальных ID (авиабилетов, ЖД билетов, MCO, заказов) из описания
    """
    if not text:
        return set()
        
    found = set()
    
    # 1. Авиабилеты: 2-4 символа + дефис + 10 цифр (555-2396327717, 15К-6111740577)
    air_tickets = re.findall(r'\b[0-9А-ЯA-Zа-яa-z]{2,4}-\d{10}\b', text)
    for t in air_tickets:
        found.add(t)
    
    # 2. Ж/Д билеты: ЭЖБ-14 цифр или 14 цифр подряд (ЭЖБ-75014939788965)
    rail_tickets = re.findall(r'\b[А-ЯA-Zа-яa-z]{3}-\d{14}\b', text)
    for t in rail_tickets:
        found.add(t)
        found.add(t.split('-')[-1])
        
    rail_digits = re.findall(r'\b75\d{12}\b', text)
    for d in rail_digits:
        found.add(d)
        found.add(f"ЭЖБ-{d}")
        
    # 3. Заказы отелей: 7-10 цифр после Заказ/№
    orders = re.findall(r'(?:Заказ|заказ|заказа|№)\s*(\d{7,10})\b', text)
    for o in orders:
        found.add(o)
        
    return found

def get_primary_id(ids: Set[str]) -> str:
    """
    Возвращает наиболее представительный сквозной ID для отображения в отчете
    """
    if not ids:
        return "N/A"
    with_hyphen = [i for i in ids if '-' in i]
    if with_hyphen:
        return with_hyphen[0]
    sorted_ids = sorted(list(ids), key=lambda x: len(x), reverse=True)
    return sorted_ids[0]

