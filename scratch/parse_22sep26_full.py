import os
import re
import json
import fitz
import pytesseract
from PIL import Image
from collections import defaultdict

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\qchen\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

pdf_path = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\22SEP26.pdf'
doc = fitz.open(pdf_path)

PAGE_CONFIG = {
    4:  {"inst": "2222", "media": "TSB", "type": "DI"},
    5:  {"inst": "2222", "media": "FTM", "type": "DI"},
    6:  {"inst": "2222", "media": "TSB", "type": "MF+DMEM"},
    7:  {"inst": "2222", "media": "FTM", "type": "MF+DMEM"},
    8:  {"inst": "2222", "media": "TSB", "type": "PBS"},
    9:  {"inst": "2222", "media": "FTM", "type": "PBS"},
    10: {"inst": "2222", "media": "TSB", "type": "DI"},
    11: {"inst": "2222", "media": "FTM", "type": "DI"},
    15: {"inst": "2011", "media": "TSB", "type": "MF"},
    16: {"inst": "2011", "media": "FTM", "type": "MF"},
    17: {"inst": "2011", "media": "TSB", "type": "DI"},
    18: {"inst": "2011", "media": "FTM", "type": "DI"},
}

DAILY_ATP = {
    "2222": 103791,
    "2011": 97310
}

def clean_etx_id(txt):
    m = re.search(r"ET[X|K|R][- ]?(\d{6})[- ]?(\d{4})", txt, re.I)
    if not m:
        return None
    d1, d2 = m.group(1), m.group(2)
    if d1.startswith("20"):
        d1 = "26" + d1[2:]
    if d1 == "260014": d1 = "260914"
    if d1 == "260011": d1 = "260911"
    if d1 == "260012": d1 = "260912"
    if d1 == "260010": d1 = "260910"
    return f"ETX-{d1}-{d2}"

sample_data = defaultdict(lambda: {
    "TSB_readings": [],
    "FTM_readings": [],
    "CVs": [],
    "instruments": set(),
    "pages": set(),
    "results": [],
    "raw_rows": []
})

all_rows_debug = []

for pno, cfg in sorted(PAGE_CONFIG.items()):
    page = doc[pno - 1]
    pix = page.get_pixmap(dpi=200)
    img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
    
    # Let's OCR the table area (approx from 15% down to 88%)
    w, h = img.size
    table_crop = img.crop((0, int(h * 0.15), w, int(h * 0.88)))
    txt = pytesseract.image_to_string(table_crop, config="--psm 6")
    
    lines = [l.strip() for l in txt.split("\n") if l.strip()]
    print(f"Page {pno:02d} ({cfg['inst']} {cfg['media']} {cfg['type']}): {len(lines)} lines")
    
    for line in lines:
        all_rows_debug.append({"page": pno, "inst": cfg["inst"], "media": cfg["media"], "text": line})
        
        # Check if control line
        if any(c in line.lower() for c in ["control", "blank", "cal ok"]):
            continue
            
        etx = clean_etx_id(line)
        if not etx:
            continue
            
        sample_data[etx]["instruments"].add(cfg["inst"])
        sample_data[etx]["pages"].add(pno)
        sample_data[etx]["raw_rows"].append({"page": pno, "media": cfg["media"], "text": line})
        
        # Result determination
        res = "Negative"
        if "overload" in line.lower():
            res = "Overload"
        elif "positive" in line.lower():
            res = "Positive"
        sample_data[etx]["results"].append(res)
        
        # Try to parse RLU
        # Pattern 1: RLU1 | RLU2 | RLU | Result
        # e.g., "1700] 1752 | 1726 | Negative"
        m_r = re.search(r"(\d{2,7})\s*[|;\]}\s]+(\d{2,7})\s*[|;\]}\s]+(\d{2,7})\s*[|;\]}\s]*(?:Negative|Positive|Cal OK|Overload)", line, re.I)
        if m_r:
            r1, r2, rlu = int(m_r.group(1)), int(m_r.group(2)), int(m_r.group(3))
            if cfg["media"] == "TSB":
                sample_data[etx]["TSB_readings"].append(rlu)
            elif cfg["media"] == "FTM":
                sample_data[etx]["FTM_readings"].append(rlu)
        else:
            # Look for number right before Negative/Positive
            m_before = re.search(r"(\d{2,7})\s*[|;\]}\s]*(?:Negative|Positive|Cal OK|Overload)", line, re.I)
            if m_before:
                rlu = int(m_before.group(1))
                if cfg["media"] == "TSB":
                    sample_data[etx]["TSB_readings"].append(rlu)
                elif cfg["media"] == "FTM":
                    sample_data[etx]["FTM_readings"].append(rlu)
            else:
                # If numbers are at end or after result
                nums = re.findall(r"\b\d{2,7}\b", line)
                if nums:
                    # Usually the largest or last number before date
                    pass
                    
        # Extract CV%
        # Pattern: after AM/PM or near end of line
        m_cv = re.search(r"\b(?:AM|PM)\b[|;\]}\s]+(\d{1,2})\b", line, re.I)
        if m_cv:
            cv_val = int(m_cv.group(1))
            sample_data[etx]["CVs"].append(cv_val)
        else:
            m_cv2 = re.search(r"(?:Negative|Positive)\b[^\n\r]*?\b(\d{1,2})\s+\d{1,2}\s*$", line, re.I)
            if m_cv2:
                sample_data[etx]["CVs"].append(int(m_cv2.group(1)))

print(f"\nTotal unique ETX samples found: {len(sample_data)}")

# Master database
records = []
for sid in sorted(sample_data.keys()):
    info = sample_data[sid]
    insts = list(info["instruments"])
    inst = insts[0] if insts else "UNKNOWN"
    atp = DAILY_ATP.get(inst, "UNKNOWN")
    
    tsb_max = max(info["TSB_readings"]) if info["TSB_readings"] else None
    ftm_max = max(info["FTM_readings"]) if info["FTM_readings"] else None
    cv_max = max(info["CVs"]) if info["CVs"] else None
    
    is_overload = "Overload" in info["results"]
    is_pos = "Positive" in info["results"]
    
    status = "PASS"
    notes = "100% Valid"
    if is_overload:
        status = "BLOCK (Overload)"
        notes = "Overload detected"
    elif is_pos:
        status = "BLOCK (Positive)"
        notes = "Positive detected"
    elif cv_max is not None and cv_max >= 30:
        status = f"BLOCK (CV {cv_max}% >= 30%)"
        notes = f"CV {cv_max}% >= 30%"
        
    rec = {
        "Sample": sid,
        "Instrument": inst,
        "Daily ATP": atp,
        "Max TSB RLU": tsb_max,
        "Max FTM RLU": ftm_max,
        "Max CV%": f"{cv_max}%" if cv_max is not None else "N/A",
        "Pages": sorted(list(info["pages"])),
        "Status": status,
        "Notes": notes,
        "Raw Rows": info["raw_rows"]
    }
    records.append(rec)

os.makedirs("scratch", exist_ok=True)
with open(r"scratch\celsis_22sep26_database.json", "w", encoding="utf-8") as f:
    json.dump(records, f, indent=2)

with open(r"scratch\celsis_22sep26_debug_rows.json", "w", encoding="utf-8") as f:
    json.dump(all_rows_debug, f, indent=2)

print("Saved scratch/celsis_22sep26_database.json successfully!")
