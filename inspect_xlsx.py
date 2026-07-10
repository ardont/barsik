import openpyxl
import sys

file_path = r"C:\Users\Maxim\Downloads\08.07_1.xlsx"
out_path = "sheet_inspection.txt"

with open(out_path, "w", encoding="utf-8") as f:
    f.write("=== INSPECTING EXCEL FILE ===\n")
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        f.write(f"Sheets in workbook: {wb.sheetnames}\n\n")
        
        for name in wb.sheetnames:
            ws = wb[name]
            f.write(f"--- Sheet: {name} ---\n")
            f.write(f"Dimensions: {ws.dimensions}\n")
            f.write("First 15 rows:\n")
            for r in range(1, 26):
                row_vals = [ws.cell(row=r, column=c).value for c in range(1, 15)]
                # Format non-None values to string, None to empty
                row_str = " | ".join([str(val) if val is not None else "" for val in row_vals])
                f.write(f"Row {r:02d}: {row_str}\n")
            f.write("\n")
        wb.close()
        f.write("Inspection completed successfully.\n")
    except Exception as e:
        import traceback
        f.write(f"Error occurred: {str(e)}\n")
        traceback.print_exc(file=f)

print("Saved inspection to", out_path)
