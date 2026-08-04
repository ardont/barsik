import openpyxl
from engine.loader import parse_bt_sheet
from config import COL_MAP_SINGLE_BT

file_path = r"C:\Users\Maxim\Downloads\20-26_07_26.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

bt_raw = parse_bt_sheet(ws, COL_MAP_SINGLE_BT)

for item in bt_raw:
    if 60 <= item.row <= 75:
        print(f"Item Row {item.row:2d} | Doc: {repr(item.doc[:20])} | Amt: {item.amount} | Desc: {repr(item.desc[:40])}")
