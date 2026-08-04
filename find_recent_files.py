import os
import datetime

downloads_dir = r"C:\Users\Maxim\Downloads"
for name in os.listdir(downloads_dir):
    if name.endswith(".xlsx"):
        path = os.path.join(downloads_dir, name)
        mtime = os.path.getmtime(path)
        mdate = datetime.datetime.fromtimestamp(mtime)
        # Check if modified today, July 21, 2026
        if mdate.year == 2026 and mdate.month == 7 and mdate.day == 21:
            print(f"File: {name}")
            print(f"  Modified: {mdate}")
            print(f"  Size: {os.path.getsize(path)} bytes")
