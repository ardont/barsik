# -*- coding: utf-8 -*-
"""
Модуль загрузки и первичного парсинга данных из Excel
"""

import openpyxl
import re
from collections import defaultdict
from typing import List, Tuple, Any, Optional
from models import ServiceItem
from engine.normalizer import clean_text, classify_service, extract_identifiers

def to_float(val: Any) -> Optional[float]:
    """
    Безопасное приведение значения к float с поддержкой запятых, пробелов и текстового мусора
    """
    if val is None or val == "":
        return None
    try:
        s = str(val).replace(" ", "").replace("\xa0", "").replace(",", ".").strip()
        # Извлекаем первое число, если строка начинается с цифры или знака
        match = re.match(r'^[-+]?\d*\.?\d+', s)
        if match:
            return float(match.group(0))
        return float(s)
    except (ValueError, TypeError):
        return None

def load_data(file_path: str) -> Tuple[List[ServiceItem], List[ServiceItem]]:
    """
    Загружает файл Excel и разбирает его на два потока:
    - TicketProf (Колонки A-D)
    - Bars Tour (Колонки H-M)
    Выполняет аллокацию отрицательных сумм корректировок.
    """
    import re # Нужен для регулярки в to_float
    
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    # Ищем лист "Лист2", иначе берем первый активный
    if "Лист2" in wb.sheetnames:
        ws = wb["Лист2"]
    else:
        ws = wb.active
        
    tp_items: List[ServiceItem] = []
    bt_items: List[ServiceItem] = []
    
    # Вспомогательные переменные для отслеживания текущих заголовков документов
    current_doc_tp = None
    current_date_tp = None
    current_doc_tp_row = None
    current_doc_tp_amt = 0.0
    
    current_doc_bt = None
    current_date_bt = None
    current_doc_bt_row = None
    current_doc_bt_amt = 0.0
    
    # Карта сумм документов TicketProf по ключу (date, doc)
    doc_sums = {}
    
    # Построчный разбор
    for r in range(2, ws.max_row + 1):
        # --- Разбор левой стороны (TicketProf) ---
        val_a = ws.cell(row=r, column=1).value
        val_b = ws.cell(row=r, column=2).value
        val_c = ws.cell(row=r, column=3).value
        val_d = ws.cell(row=r, column=4).value
        
        is_header_tp = False
        if val_a is not None and val_b is not None:
            if any(kw in str(val_b) for kw in ["Продажа", "Оплата", "Возврат", "Корректировка"]):
                is_header_tp = True
                
        if is_header_tp:
            current_date_tp = str(val_a).strip()
            current_doc_tp = str(val_b).strip()
            current_doc_tp_row = r
            
            c_val = to_float(val_c)
            d_val = to_float(val_d)
            current_doc_tp_amt = c_val if c_val is not None else (-d_val if d_val is not None else 0.0)
            doc_sums[(current_date_tp, current_doc_tp)] = current_doc_tp_amt
        elif val_a is not None and val_b is None:
            desc = str(val_a).strip()
            if not any(kw in desc for kw in ["Обороты за период", "Сальдо конечное", "Итого"]):
                c_val = to_float(val_c)
                d_val = to_float(val_d)
                amt = c_val if c_val is not None else (-d_val if d_val is not None else 0.0)
                tp_items.append(ServiceItem(
                    row=r,
                    date=current_date_tp or "",
                    doc=current_doc_tp or "",
                    desc=desc,
                    clean_desc=clean_text(desc),
                    service_type=classify_service(desc),
                    amount=amt,
                    allocated_amount=amt,
                    ids=extract_identifiers(desc),
                    source="TP"
                ))
                
        # --- Разбор правой стороны (Bars Tour) ---
        val_h = ws.cell(row=r, column=8).value
        val_i = ws.cell(row=r, column=9).value
        val_j = ws.cell(row=r, column=10).value
        val_l = ws.cell(row=r, column=12).value
        val_m = ws.cell(row=r, column=13).value
        
        is_header_bt = False
        if val_h is not None and val_i is not None:
            if any(kw in str(val_i) for kw in ["Приход", "Оплата", "Возврат", "Принято"]):
                is_header_bt = True
                
        if is_header_bt:
            current_date_bt = str(val_h).strip()
            current_doc_bt = str(val_i).strip()
            current_doc_bt_row = r
            j_val = to_float(val_j)
            current_doc_bt_amt = j_val if j_val is not None else 0.0
        elif val_h is not None and val_i is None:
            desc = str(val_h).strip()
            if not any(kw in desc for kw in ["Обороты за период", "Сальдо конечное", "Итого"]):
                j_val = to_float(val_j)
                amt = j_val if j_val is not None else 0.0
                # Пропускаем строки с нулевой суммой и пустым ID, чтобы не перегружать таблицу
                if amt == 0.0 and not extract_identifiers(desc):
                    continue
                
                # Безопасное чтение прибыли и нетто
                profit_val = to_float(val_l)
                net_val = to_float(val_m)
                
                bt_items.append(ServiceItem(
                    row=r,
                    date=current_date_bt or "",
                    doc=current_doc_bt or "",
                    desc=desc,
                    clean_desc=clean_text(desc),
                    service_type=classify_service(desc),
                    amount=amt,
                    allocated_amount=amt,
                    ids=extract_identifiers(desc),
                    profit=profit_val,
                    net=net_val,
                    source="BT"
                ))
                
    wb.close()
    
    # --- Алгоритм распределения (аллокации) отрицательных сумм корректировок в TicketProf ---
    tp_groups = defaultdict(list)
    for item in tp_items:
        key = (item.date, item.doc)
        tp_groups[key].append(item)
            
    # Распределяем
    for key, items in tp_groups.items():
        doc_amt = doc_sums.get(key, 0.0)
        if doc_amt < 0:
            # Находим элементы с нулевой стоимостью
            zero_items = [item for item in items if item.amount == 0.0]
            non_zero_sum = sum(item.amount for item in items if item.amount != 0.0)
            remaining_to_allocate = doc_amt - non_zero_sum
            
            if zero_items and remaining_to_allocate != 0.0:
                allocated = 0.0
                for item in zero_items:
                    # Ищем соответствующий приход BT с отрицательной суммой по пересечению ID
                    bt_match = None
                    for bt in bt_items:
                        if bt.ids.intersection(item.ids) and bt.service_type == item.service_type and bt.amount < 0:
                            bt_match = bt
                            break
                    if bt_match:
                        item.allocated_amount = bt_match.amount
                        allocated += bt_match.amount
                    else:
                        # Делим поровну
                        item.allocated_amount = remaining_to_allocate / len(zero_items)
                        
                # Корректируем погрешность округления на последнем элементе
                total_allocated = sum(item.allocated_amount for item in zero_items)
                if abs(total_allocated - remaining_to_allocate) > 0.01:
                    zero_items[-1].allocated_amount += (remaining_to_allocate - total_allocated)
                    
    return tp_items, bt_items
