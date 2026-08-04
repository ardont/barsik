import openpyxl
from config import COL_MAP_SINGLE_BT
from engine.normalizer import clean_text, classify_service, extract_identifiers

file_path = r"C:\Users\Maxim\Downloads\20-26_07_26.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

col_map = COL_MAP_SINGLE_BT

for r in [68, 70]:
    val_date = ws.cell(row=r, column=col_map["date"]).value
    desc = str(val_date).strip()
    cdesc = clean_text(desc)
    stype = classify_service(desc)
    ids = extract_identifiers(desc)
    print(f"Row {r:2d}:")
    print(f"  desc: {repr(desc)}")
    print(f"  clean_desc: {repr(cdesc)}")
    print(f"  stype: {repr(stype)}")
    print(f"  ids: {ids}")
