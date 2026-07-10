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
