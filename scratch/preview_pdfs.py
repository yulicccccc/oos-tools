import os
import glob

scratch_dir = r"C:\Users\qchen\.gemini\antigravity\brain\f2811f7a-6e22-4e06-be5b-f4e86a5e371a\scratch"
txt_files = sorted(glob.glob(os.path.join(scratch_dir, "pdf_*_extracted.txt")))

print(f"Found {len(txt_files)} text files.")

for fpath in txt_files:
    fname = os.path.basename(fpath)
    print(f"\n==================== {fname} ====================")
    with open(fpath, "r", encoding="utf-8") as f:
        lines = f.readlines()
    print("--- First 30 lines ---")
    for line in lines[:30]:
        print(line.rstrip())
