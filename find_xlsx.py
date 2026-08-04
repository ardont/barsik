import os
from pathlib import Path

workspace_dir = Path(r"c:\Users\Maxim\Desktop\фокусы\инструменты\тиектпрофи\барсик")
for path in workspace_dir.glob("**/*.xlsx"):
    print("Found XLSX:", path)
