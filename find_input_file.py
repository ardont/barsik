import os
import openpyxl

downloads_dir = r"C:\Users\Maxim\Downloads"
for name in os.listdir(downloads_dir):
    if name.endswith(".xlsx") and not name.endswith("_сопоставлено.xlsx"):
        path = os.path.join(downloads_dir, name)
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            print(f"File: {name}")
            print(f"  Sheets: {wb.sheetnames}")
            wb.close()
        except Exception as e:
            print(f"Error reading {name}: {e}")
