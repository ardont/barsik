import openpyxl
from engine.loader import parse_bt_sheet
from config import COL_MAP_SINGLE_BT

file_path = r"C:\Users\Maxim\Downloads\20-26_07_26.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

bt_raw = parse_bt_sheet(ws, COL_MAP_SINGLE_BT)

print("RAW items before grouping:")
for item in bt_raw:
    if "555-2397261625" in str(item.ids) and item.service_type == "Flight":
        print(f"  Row {item.row:2d} | Amt: {item.amount} | clean_desc: {repr(item.clean_desc)}")

from collections import defaultdict
grouped_bt = defaultdict(list)
for item in bt_raw:
    key = (item.doc, item.clean_desc, item.service_type, frozenset(item.ids))
    grouped_bt[key].append(item)

print("\nGROUPED keys matching 555-2397261625 and Flight:")
for key, items in grouped_bt.items():
    if "555-2397261625" in str(key) and key[2] == "Flight":
        print(f"  Key: {key}")
        print(f"  Items count: {len(items)}")
        for x in items:
            print(f"    Row {x.row:2d} | Amt: {x.amount}")
