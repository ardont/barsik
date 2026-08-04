import openpyxl

from engine.loader import load_data
from engine.matcher import match_records

print("--- Step 1: Loading raw file 13-19.07.xlsx ---")
tp_items, bt_items = load_data(r'C:\Users\Maxim\Downloads\13-19.07.xlsx')

print(f"\nLoaded TicketProf items: {len(tp_items)}")
for item in tp_items:
    print(f"  TP [row {item.row:2d}]: {item.doc} | {item.desc} | amt={item.amount} | ids={item.ids} | type={item.service_type}")

print(f"\nLoaded BarsTour items: {len(bt_items)}")
for item in bt_items:
    print(f"  BT [row {item.row:2d}]: {item.doc} | {item.desc} | amt={item.amount} | ids={item.ids} | type={item.service_type}")

print("\n--- Step 2: Running Matching Engine ---")
matches, unmatched_tp, unmatched_bt = match_records(tp_items, bt_items, manual_links={}, settings={'simple_mode': True})

print(f"\nMatches count: {len(matches)}")
for tp, bt, method, score in matches:
    print(f" MATCH [{method:20s}] TP(row {tp.row}, amt={tp.amount}, ids={tp.ids}, desc={tp.desc}) <-> BT(row {bt.row}, amt={bt.amount}, ids={bt.ids}, desc={bt.desc})")

print(f"\nUnmatched TP count: {len(unmatched_tp)}")
for r in unmatched_tp:
    print(f" UNMATCHED TP [row {r.row:2d}]: {r.doc} | {r.desc} | amt={r.amount} | ids={r.ids}")

print(f"\nUnmatched BT count: {len(unmatched_bt)}")
for r in unmatched_bt:
    print(f" UNMATCHED BT [row {r.row:2d}]: {r.doc} | {r.desc} | amt={r.amount} | ids={r.ids}")
