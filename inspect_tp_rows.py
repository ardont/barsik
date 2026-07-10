import openpyxl

file_path = r"C:\Users\Maxim\Downloads\08.07_1.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb["Лист2"]

print("=== CHECKING COLUMNS A-G FOR ROWS 21 TO 30 ===")
for r in range(21, 31):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 8)]
    print(f"Row {r:3d}: {vals}")
wb.close()
