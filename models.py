# -*- coding: utf-8 -*-
"""
Модели данных приложения
"""

from dataclasses import dataclass, field
from typing import Set, Optional, List, Dict

@dataclass
class ServiceItem:
    row: int
    date: str
    doc: str
    desc: str
    clean_desc: str
    service_type: str
    amount: float
    allocated_amount: float
    ids: Set[str] = field(default_factory=set)
    matched: bool = False
    source: str = "TP"  # "TP" (TicketProf) или "BT" (Bars Tour)
    
    # Специфичные для Bars Tour поля
    profit: Optional[float] = None
    net: Optional[float] = None
    
    # Ссылка на сопоставленный элемент
    matched_row: Optional[int] = None
    match_method: Optional[str] = None
    match_score: float = 0.0

    def get_status_text(self, other: Optional['ServiceItem'] = None) -> str:
        """
        Возвращает русскоязычный статус сопоставления
        """
        if not self.matched:
            if self.source == "TP":
                return "В Тикете, нет в Барсе"
            else:
                if self.service_type == "Fee":
                    return "Норма (Сбор в БТ)"
                return "В Барсе, нет в Тикете"
        
        if other is not None:
            # Сравниваем суммы
            tp_amt = self.allocated_amount if self.source == "TP" else other.allocated_amount
            bt_amt = other.amount if self.source == "TP" else self.amount
            profit = bt_amt - tp_amt
            
            # Логика для отелей: прибыль должна соответствовать настроенному проценту маржи (только для продаж, т.е. bt_amt > 0)
            if (self.service_type == "Hotel" or other.service_type == "Hotel"):
                if bt_amt > 0:
                    margin_pct = getattr(ServiceItem, 'hotel_margin', 10.0)
                    expected_profit = (margin_pct / 100.0) * bt_amt
                    if abs(profit - expected_profit) > 0.01:
                        return "Нетипичная маржа"
                else:
                    # При возврате/корректировке отеля прибыль должна быть 0.0
                    if abs(profit) > 0.01:
                        return "Несовпадение по суммам"
            else:
                # Для остальных услуг, если прибыль отрицательная - это ошибка (убыток)
                if profit < 0:
                    return "Несовпадение по суммам"
                    
        return "Совпадение"

@dataclass
class ReconciliationSummary:
    total_tp_count: int = 0
    total_bt_count: int = 0
    total_tp_sum: float = 0.0
    total_bt_sum: float = 0.0
    
    matched_tp_count: int = 0
    matched_bt_count: int = 0
    matched_tp_sum: float = 0.0
    matched_bt_sum: float = 0.0
    
    unmatched_tp_count: int = 0
    unmatched_bt_count: int = 0
    unmatched_tp_sum: float = 0.0
    unmatched_bt_sum: float = 0.0
    
    total_profit: float = 0.0
