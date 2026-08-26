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
    narratives = []
    for b in blocks:
        # Ignore signatures
        if "/SubFilter" in b or "/Contents" in b:
            continue
        lines = [l.strip() for l in b.strip().split("\n") if l.strip()]
        if len(lines) >= 2:
            field_name = lines[0]
            val = "\n".join(lines[1:])
            if any(k in val for k in ["Phase I", "investigation", "monitoring", "CFU", "BSC", "plate", "colony", "growth", "SOP"]):
                narratives.append((field_name, val))
    
    print(f"Found {len(narratives)} relevant narrative fields:")
    for fn, val in narratives:
        print(f"\n--- [{fn}] ---")
        print(val[:500] + ("..." if len(val) > 500 else ""))
