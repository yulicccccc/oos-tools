import os
import glob

scratch_dir = r"C:\Users\qchen\.gemini\antigravity\brain\f2811f7a-6e22-4e06-be5b-f4e86a5e371a\scratch"
files = sorted(glob.glob(os.path.join(scratch_dir, "pdf_*_fields.txt")))

for fpath in files:
    fname = os.path.basename(fpath)
    print(f"\n==================== {fname} ====================")
    with open(fpath, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Print first few fields
    blocks = content.split("----------------------------------------")
    print(f"Total populated blocks: {len(blocks)}")
    for b in blocks:
        if any(keyword in b for keyword in ["Test Name:", "Environmental Monitoring", "Phase I", "Summary", "smart_", "Field: 'Text Field"]):
            lines = [line.strip() for line in b.strip().split("\n") if line.strip()]
            if len(lines) > 1:
                field_line = lines[0]
                val_snippet = " | ".join(lines[1:])[:200]
                print(f"  {field_line} => {val_snippet}")
