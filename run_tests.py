# -*- coding: utf-8 -*-
"""
Автоматизированный скрипт тестирования сверки и генерации отчетов
"""

import os
from engine.loader import load_data
from engine.matcher import match_records
from engine.calculator import calculate_reconciliation
from reports.excel_export import export_to_excel
from reports.word_export import export_to_word

def run_tests():
    print("=== ЗАПУСК ТЕСТОВ СВЕРКИ ===")
    
    file_path = r"C:\Users\Maxim\Downloads\08.07_1.xlsx"
    if not os.path.exists(file_path):
        print(f"Ошибка: Исходный файл не найден по пути: {file_path}")
        return
        
    print(f"Загрузка файла: {file_path}...")
    tp_items, bt_items = load_data(file_path)
    print(f"Загружено записей TicketProf: {len(tp_items)}")
    print(f"Загружено записей Bars Tour: {len(bt_items)}")
    
    assert len(tp_items) > 0, "Количество записей TicketProf должно быть больше 0"
    assert len(bt_items) > 0, "Количество записей Bars Tour должно быть больше 0"
    
    print("\nЗапуск алгоритма сопоставления...")
    manual_links = {}  # Для тестов используем пустой кэш
    matches, unmatched_tp, unmatched_bt = match_records(tp_items, bt_items, manual_links)
    
    print(f"Успешно сопоставлено пар: {len(matches)}")
    print(f"Осталось несопоставленных TicketProf: {len(unmatched_tp)}")
    print(f"Осталось несопоставленных Bars Tour: {len(unmatched_bt)}")
    
    print("\nРасчет финансовых показателей...")
    summary = calculate_reconciliation(tp_items, bt_items, matches)
    print(f"Общая сумма продаж TicketProf: {summary.total_tp_sum:,.2f} руб.")
    print(f"Общая сумма приходов Bars Tour: {summary.total_bt_sum:,.2f} руб.")
    print(f"Итоговая прибыль: {summary.total_profit:,.2f} руб.")
    
    print("\nПроверка экспорта в Excel...")
    excel_out = "test_report.xlsx"
    export_to_excel(tp_items, bt_items, matches, unmatched_tp, unmatched_bt, summary, excel_out)
    assert os.path.exists(excel_out), "Файл Excel отчета должен быть создан"
    print(f"Excel отчет успешно сохранен в: {os.path.abspath(excel_out)}")
    os.remove(excel_out) # удаляем временный файл
    
    print("\nПроверка экспорта в Word...")
    word_out = "test_reconciliation_act.docx"
    export_to_word(unmatched_tp, unmatched_bt, summary, word_out)
    assert os.path.exists(word_out), "Файл Word Акта сверки должен быть создан"
    print(f"Word Акт сверки успешно сохранен в: {os.path.abspath(word_out)}")
    os.remove(word_out) # удаляем временный файл
    
    print("\n=== ВСЕ ТЕСТЫ ПРОЙДЕНЫ УСПЕШНО! ===")

if __name__ == "__main__":
    try:
        run_tests()
    except Exception as e:
        import traceback
        print(f"\nОшибка во время выполнения тестов: {e}")
        traceback.print_exc()
