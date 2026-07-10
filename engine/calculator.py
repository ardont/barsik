# -*- coding: utf-8 -*-
"""
Модуль финансовых расчетов и анализа
"""

from typing import List, Tuple
from models import ServiceItem, ReconciliationSummary

def calculate_reconciliation(
    tp_items: List[ServiceItem],
    bt_items: List[ServiceItem],
    matches: List[Tuple[ServiceItem, ServiceItem, str, float]]
) -> ReconciliationSummary:
    """
    Рассчитывает общую статистику сверки и маржинальность
    """
    summary = ReconciliationSummary()
    
    # 1. Считаем общие суммы и количество по системам
    summary.total_tp_count = len(tp_items)
    summary.total_tp_sum = sum(item.allocated_amount for item in tp_items)
    
    summary.total_bt_count = len(bt_items)
    summary.total_bt_sum = sum(item.amount for item in bt_items)
    
    # 2. Считаем суммы сопоставленных позиций
    # Множества уникальных номеров строк для исключения дубликатов при суммировании
    matched_tp_rows = set()
    matched_bt_rows = set()
    
    for tp_item, bt_item, _, _ in matches:
        matched_tp_rows.add(tp_item.row)
        matched_bt_rows.add(bt_item.row)
        
        # Рассчитываем прибыль для каждой сопоставленной пары
        # Формула: Прибыль = Bars Tour (Gross) - TicketProf (Net)
        profit = bt_item.amount - tp_item.allocated_amount
        tp_item.profit = profit
        bt_item.profit = profit
        
        tp_item.net = tp_item.allocated_amount
        bt_item.net = tp_item.allocated_amount
        
        summary.total_profit += profit
        
    summary.matched_tp_count = len(matched_tp_rows)
    summary.matched_tp_sum = sum(item.allocated_amount for item in tp_items if item.row in matched_tp_rows)
    
    summary.matched_bt_count = len(matched_bt_rows)
    summary.matched_bt_sum = sum(item.amount for item in bt_items if item.row in matched_bt_rows)
    
    # 3. Считаем несопоставленные позиции
    summary.unmatched_tp_count = summary.total_tp_count - summary.matched_tp_count
    summary.unmatched_tp_sum = sum(item.allocated_amount for item in tp_items if not item.matched)
    
    summary.unmatched_bt_count = summary.total_bt_count - summary.matched_bt_count
    summary.unmatched_bt_sum = sum(item.amount for item in bt_items if not item.matched)
    
    return summary
