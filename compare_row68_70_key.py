import openpyxl
from engine.loader import parse_bt_sheet
from config import COL_MAP_SINGLE_BT

file_path = r"C:\Users\Maxim\Downloads\20-26_07_26.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

bt_items = parse_bt_sheet(ws, COL_MAP_SINGLE_BT)

r68 = [x for x in bt_items if x.row == 68][0]
r70 = [x for x in bt_items if x.row == 70][0]

key68 = (r68.doc, r68.clean_desc, r68.service_type, frozenset(r68.ids))
key70 = (r70.doc, r70.clean_desc, r70.service_type, frozenset(r70.ids))

print("Key 68:", key68)
print("Key 70:", key70)
print("Equal?:", key68 == key70)
if key68 != key70:
    for idx, (a, b) in enumerate(zip(key68, key70)):
        if a != b:
            print(f"Diff at index {idx}: {repr(a)} != {repr(b)}")
