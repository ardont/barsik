# -*- coding: utf-8 -*-
"""
Модуль интеллектуального сопоставления данных
"""

import difflib
from typing import List, Tuple, Dict, Set
from models import ServiceItem

def get_compare_amount(b: ServiceItem, service_type: str) -> float:
    if service_type == "Hotel" and b.net is not None:
        return b.net
    return b.amount

def match_records(
    tp_items: List[ServiceItem], 
    bt_items: List[ServiceItem], 
    manual_links: Dict[str, str],
    settings: dict = None
) -> Tuple[List[Tuple[ServiceItem, ServiceItem, str, float]], List[ServiceItem], List[ServiceItem]]:
    """
    Выполняет сопоставление элементов двух систем в едином стандартном режиме:
    1. По кэшу ручных связей (manual_links).
    2. По пересечению уникальных ID (билетов/заказов/MCO) и типу услуги.
    3. По точному совпадению очищенного наименования и типу услуги.
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
            bt_matches = [b for b in bt_items if b.clean_desc == target_clean_bt and not b.matched and b.service_type == tp.service_type and ((tp.allocated_amount >= 0 and b.amount >= 0) or (tp.allocated_amount < 0 and b.amount < 0))]
            if bt_matches:
                best_bt = min(bt_matches, key=lambda b: abs(get_compare_amount(b, tp.service_type) - tp.allocated_amount))
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
            bt_matches = [b for b in bt_items if b.ids.intersection(tp.ids) and b.service_type == tp.service_type and not b.matched and ((tp.allocated_amount >= 0 and b.amount >= 0) or (tp.allocated_amount < 0 and b.amount < 0))]
            if bt_matches:
                best_bt = min(bt_matches, key=lambda b: abs(get_compare_amount(b, tp.service_type) - tp.allocated_amount))
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
            bt_matches = [b for b in bt_items if b.clean_desc == tp.clean_desc and b.service_type == tp.service_type and not b.matched and ((tp.allocated_amount >= 0 and b.amount >= 0) or (tp.allocated_amount < 0 and b.amount < 0))]
            if bt_matches:
                best_bt = min(bt_matches, key=lambda b: abs(get_compare_amount(b, tp.service_type) - tp.allocated_amount))
                tp.matched = True
                best_bt.matched = True
                tp.matched_row = best_bt.row
                best_bt.matched_row = tp.row
                tp.match_method = "Точное имя"
                best_bt.match_method = "Точное имя"
                tp.match_score = 1.0
                best_bt.match_score = 1.0
                matches.append((tp, best_bt, "Точное имя", 1.0))
                
    # Собираем нераспределенные списки
    unmatched_tp = [tp for tp in tp_items if not tp.matched]
    unmatched_bt = [bt for bt in bt_items if not bt.matched]
    
    return matches, unmatched_tp, unmatched_bt
