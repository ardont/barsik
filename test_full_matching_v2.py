import openpyxl
import re
from dataclasses import dataclass, field
from typing import Set, Optional, List, Dict, Tuple
from collections import defaultdict

# --- Normalized Regexes & Functions ---

def clean_text(text: str) -> str:
    if not text:
        return ""
    t = text.lower()
    t = re.sub(r'\b\d{2}\.\d{2}\.(?:\d{4}|\d{2})\b', '', t)
    t = re.sub(r'\bс\s+\d{2}\.\d{2}\s+по\s+\d{2}\.\d{2}\b', '', t)
    t = re.sub(r'\bс\s+\d{2}\.\d{2}\.\d{2,4}\s+по\s+\d{2}\.\d{2}\.\d{2,4}\b', '', t)
    # Remove passenger info / ticket info in brackets
    t = re.sub(r'\([a-z0-9\s\/\-\:\.\,№а-я]+\)', '', t)
    t = re.sub(r'\b[a-z]+/[a-z]+\b', '', t)
    t = re.sub(r'\b[0-9а-яa-z]{2,4}-\d{10,14}\b', '', t)
    t = re.sub(r'\b\d{13,14}\b', '', t)
    t = re.sub(r'\b(?:заказ|заказа|№)?\s*\d+\b', '', t)
    t = re.sub(r'[^\w\s]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def classify_service(text: str) -> str:
    if not text:
        return "Other"
    t = text.lower()
    if any(kw in t for kw in ["проживание", "отель", "г-ца", "гостиница"]):
        return "Hotel"
    if any(kw in t for kw in ["выбор места", "место", "mco"]):
        return "Seat"
    if any(kw in t for kw in ["услуга ито", "сервисный сбор", "ито"]):
        return "Fee"
    if any(kw in t for kw in ["штраф", "удержания", "удержание"]):
        return "Penalty"
    if any(kw in t for kw in ["авиабилет", "билет", "ж.д.", "перелет", "поезд"]):
        return "Flight"
    return "Other"

def extract_identifiers(text: str) -> Set[str]:
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
        # Также добавляем только цифровую часть для надежного пересечения
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
    if not ids:
        return "N/A"
    # Предчитаем префиксные билеты типа 555-..., 15К-..., ЭЖБ-...
    with_hyphen = [i for i in ids if '-' in i]
    if with_hyphen:
        return with_hyphen[0]
    # Заказы или цифры
    sorted_ids = sorted(list(ids), key=lambda x: len(x), reverse=True)
    return sorted_ids[0]

print("Done defining functions.")
