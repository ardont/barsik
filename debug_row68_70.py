import openpyxl
from engine.normalizer import clean_text, classify_service, extract_identifiers
from engine.loader import parse_bt_sheet
from config import COL_MAP_SINGLE_BT

file_path = r"C:\Users\Maxim\Downloads\20-26_07_26.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

bt_raw = parse_bt_sheet(ws, COL_MAP_SINGLE_BT)

print("Raw BT items around row 68:")
for item in bt_raw:
    if "555-2397261625" in str(item.ids):
        print(f"Row {item.row:3d} | Doc: {repr(item.doc)} | ServiceType: {repr(item.service_type)} | Amt: {item.amount} | IDs: {item.ids}")
        print(f"  desc: {repr(item.desc)}")
        print(f"  clean_desc: {repr(item.clean_desc)}")
