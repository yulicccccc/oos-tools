import os, shutil, glob
from datetime import datetime

ts = datetime.now().strftime("%Y%m%d_%H%M%S")
history_dir = ".history"
os.makedirs(history_dir, exist_ok=True)

targets = [
    "ScanRDI OOS template.docx",
    "ScanRDI OOS template 0.docx",
    "ScanRDI OOS P1 template.docx",
    "ScanRDI OOS P1 template 0.docx",
    "ScanRDI OOS template.pdf",
    "ScanRDI OOS P1 template.pdf",
]

for t in targets:
    if os.path.exists(t):
        dest = os.path.join(history_dir, f"{t}_backup_{ts}")
        shutil.copy2(t, dest)
        print(f"Backed up {t} to {dest}")
