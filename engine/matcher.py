# -*- coding: utf-8 -*-
"""
Модуль интеллектуального сопоставления данных
"""

import difflib
from typing import List, Tuple, Dict, Set
from models import ServiceItem

def match_records(
    tp_items: List[ServiceItem], 
    bt_items: List[ServiceItem], 
    manual_links: Dict[str, str]
) -> Tuple[List[Tuple[ServiceItem, ServiceItem, str, float]], List[ServiceItem], List[ServiceItem]]:
    """
    Выполняет пошаговое сопоставление элементов двух систем:
    1. По кэшу ручных связей (manual_links).
    2. По пересечению уникальных ID (билетов/заказов/MCO) и типу услуги.
    3. По точному совпадению очищенного наименования и типу услуги.
    4. По нечеткому совпадению (Fuzzy) с порогом 0.75.
    
    Возвращает:
    - Список кортежей: (TP_item, BT_item, метод_сопоставления, процент_сходства)
    - Список нераспределенных TP_items
    - Список нераспределенных BT_items
    """
    matches = []
    
    # Сбрасываем флаги сопоставления
    for tp in tp_items:
        tp.matched = False
        tp.matched_row = None
        tp.match_method = None
        tp.match_score = 0.0
    for bt in bt_items:
        bt.matched = False
        bt.matched_row = None
        bt.match_method = None
        bt.match_score = 0.0
        
    # --- Шаг 0: Ручные связи (Human-in-the-Loop) ---
    for tp in tp_items:
        if not tp.matched and tp.clean_desc in manual_links:
            target_clean_bt = manual_links[tp.clean_desc]
            # Ищем нераспределенный BT с таким же чистым описанием
            bt_matches = [b for b in bt_items if b.clean_desc == target_clean_bt and not b.matched and b.service_type == tp.service_type]
            if bt_matches:
                best_bt = min(bt_matches, key=lambda b: abs(b.amount - tp.allocated_amount))
                tp.matched = True
                best_bt.matched = True
                tp.matched_row = best_bt.row
                best_bt.matched_row = tp.row
                tp.match_method = "Ручная связь"
                best_bt.match_method = "Ручная связь"
                tp.match_score = 1.0
                best_bt.match_score = 1.0
                matches.append((tp, best_bt, "Ручная связь", 1.0))
                
    # --- Шаг 1: Сопоставление по Shared ID и типу услуги ---
    for tp in tp_items:
        if not tp.matched and tp.ids:
            # Ищем BT с пересекающимся набором ID и тем же типом услуги
            bt_matches = [b for b in bt_items if b.ids.intersection(tp.ids) and b.service_type == tp.service_type and not b.matched]
            if bt_matches:
                # Берем наиболее близкий по сумме элемент
                best_bt = min(bt_matches, key=lambda b: abs(b.amount - tp.allocated_amount))
                tp.matched = True
                best_bt.matched = True
                tp.matched_row = best_bt.row
                best_bt.matched_row = tp.row
                tp.match_method = "По ID билета/заказа"
                best_bt.match_method = "По ID билета/заказа"
                tp.match_score = 1.0
                best_bt.match_score = 1.0
                matches.append((tp, best_bt, "По ID билета/заказа", 1.0))
                
    # --- Шаг 2: Точное совпадение очищенного текста и типа услуги ---
    for tp in tp_items:
        if not tp.matched:
            bt_matches = [b for b in bt_items if b.clean_desc == tp.clean_desc and b.service_type == tp.service_type and not b.matched]
            if bt_matches:
                best_bt = min(bt_matches, key=lambda b: abs(b.amount - tp.allocated_amount))
                tp.matched = True
                best_bt.matched = True
                tp.matched_row = best_bt.row
                best_bt.matched_row = tp.row
                tp.match_method = "Точное имя"
                best_bt.match_method = "Точное имя"
                tp.match_score = 1.0
                best_bt.match_score = 1.0
                matches.append((tp, best_bt, "Точное имя", 1.0))
                
    # --- Шаг 3: Нечеткое сопоставление (Fuzzy matching) ---
    for tp in tp_items:
        if not tp.matched:
            unmatched_bt = [b for b in bt_items if not b.matched and b.service_type == tp.service_type]
            best_bt = None
            best_score = 0.0
            
            for bt in unmatched_bt:
                score = difflib.SequenceMatcher(None, tp.clean_desc, bt.clean_desc).ratio()
                if score > best_score:
                    best_score = score
                    best_bt = bt
                    
            if best_bt and best_score >= 0.75:
                tp.matched = True
                best_bt.matched = True
                tp.matched_row = best_bt.row
                best_bt.matched_row = tp.row
                tp.match_method = "Нечеткое имя"
                best_bt.match_method = "Нечеткое имя"
                tp.match_score = best_score
                best_bt.match_score = best_score
                matches.append((tp, best_bt, f"Нечеткое имя ({tp.service_type})", best_score))
                
    # Собираем нераспределенные списки
    unmatched_tp = [tp for tp in tp_items if not tp.matched]
    unmatched_bt = [bt for bt in bt_items if not bt.matched]
    
    return matches, unmatched_tp, unmatched_bt
