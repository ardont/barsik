import openpyxl
from config import COL_MAP_SINGLE_BT
from engine.loader import parse_bt_sheet

file_path = r"C:\Users\Maxim\Downloads\20-26_07_26.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

bt_raw = parse_bt_sheet(ws, COL_MAP_SINGLE_BT)

for item in bt_raw:
    if item.row in [68, 69, 70]:
        print(f"bt_raw item: row={item.row} | amt={item.amount} | type={item.service_type} | ids={item.ids}")
