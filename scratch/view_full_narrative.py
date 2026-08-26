import os
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\qchen\.gemini\antigravity\brain\f2811f7a-6e22-4e06-be5b-f4e86a5e371a\scratch"
files = sorted(glob.glob(os.path.join(scratch_dir, "pdf_*_fields.txt")))

for fpath in files:
    fname = os.path.basename(fpath)
    print(f"\n==================== {fname} ====================")
    with open(fpath, "r", encoding="utf-8") as f:
        content = f.read()
    
    blocks = content.split("----------------------------------------")
    for b in blocks:
        # Find large text blocks or summary blocks
        if len(b) > 200:
            lines = [l.strip() for l in b.strip().split("\n") if l.strip()]
            if len(lines) >= 2:
                print(f"\n>>> FIELD NAME: {lines[0]}")
                print(f"--- CONTENT ({len(b)} chars) ---")
                print("\n".join(lines[1:30]))
                if len(lines) > 30:
                    print("... [truncated] ...")
