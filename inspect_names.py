import os

downloads_dir = r"C:\Users\Maxim\Downloads"
for name in os.listdir(downloads_dir):
    if "13-19.07" in name:
        codepoints = [ord(c) for c in name]
        print(f"File: {name}")
        print(f"Codepoints: {codepoints}")
