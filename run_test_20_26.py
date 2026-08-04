import sys
import openpyxl
from engine.loader import load_data
from engine.matcher import match_records
from engine.calculator import calculate_reconciliation

file_path = r"C:\Users\Maxim\Downloads\20-26_07_26.xlsx"
tp_items, bt_items = load_data(file_path)

out = []
out.append(f"Loaded TP items: {len(tp_items)}")
out.append(f"Loaded BT items: {len(bt_items)}")

out.append("\n--- ALL TP ITEMS ---")
for item in tp_items:
    out.append(f"TP Row {item.row:3d} | Doc: {item.doc} | Desc: {item.desc[:50]}... | Amt: {item.amount} | IDs: {item.ids} | Type: {item.service_type}")

out.append("\n--- ALL BT ITEMS ---")
for item in bt_items:
    out.append(f"BT Row {item.row:3d} | Doc: {item.doc} | Desc: {item.desc[:50]}... | Amt: {item.amount} | Profit: {item.profit} | Net: {item.net} | IDs: {item.ids} | Type: {item.service_type}")

matches, unmatched_tp, unmatched_bt = match_records(tp_items, bt_items, {})
summary = calculate_reconciliation(tp_items, bt_items, matches)

out.append("\n--- MATCHES ---")
for tp, bt, method, score in matches:
    out.append(f"MATCH: TP Row {tp.row:3d} ({tp.amount:10.2f} | {tp.desc[:30]}...) <==> BT Row {bt.row:3d} ({bt.amount:10.2f} | {bt.desc[:30]}...) | Method: {method} | Profit: {bt.profit}")

out.append("\n--- UNMATCHED TP ---")
for tp in unmatched_tp:
    out.append(f"UNMATCHED TP: Row {tp.row:3d} ({tp.amount:10.2f} | {tp.desc[:40]}...) | IDs: {tp.ids} | Type: {tp.service_type}")

out.append("\n--- UNMATCHED BT ---")
for bt in unmatched_bt:
    out.append(f"UNMATCHED BT: Row {bt.row:3d} ({bt.amount:10.2f} | {bt.desc[:40]}...) | IDs: {bt.ids} | Type: {bt.service_type}")

with open("test_20_26_out_utf8.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))

print("Saved test_20_26_out_utf8.txt")
