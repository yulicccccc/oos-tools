import os
import pypdf

pdf_paths = [
    r"G:\CRO\OOS\2026\05-2026\OOS- 260422 ScanCO HS BSC1309 S3 17FEB2026 and ScanCO HS BSC1309 S1 17FEB2026 MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260469 Scan GA 1311 SettR 23FEB2026-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260541 Scan PG BSC1310 Sett2 26FEB2026-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260542 Scan DH BSC1313 SettL 26FEB2026 draft-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260543 Scan VV BSC1309 S1 27FEB2026  draft-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260361 Sterility GS BSC1314 Sett2 05FEB2026-signed by MN - KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260365 EM OA 114A cart 10FEB2026-Signed by MN - KSP.pdf"
]

out_dir = r"C:\Users\qchen\.gemini\antigravity\brain\f2811f7a-6e22-4e06-be5b-f4e86a5e371a\scratch"
os.makedirs(out_dir, exist_ok=True)

for i, path in enumerate(pdf_paths):
    print(f"--- Checking file {i+1}: {os.path.basename(path)} ---")
    if not os.path.exists(path):
        print(f"File NOT found: {path}")
        continue
    
    try:
        reader = pypdf.PdfReader(path)
        print(f"Total Pages: {len(reader.pages)}")
        text_content = []
        for p_idx, page in enumerate(reader.pages):
            txt = page.extract_text()
            text_content.append(f"=== PAGE {p_idx+1} ===\n{txt}\n")
        
        out_file = os.path.join(out_dir, f"pdf_{i+1}_extracted.txt")
        with open(out_file, "w", encoding="utf-8") as f:
            f.write("\n".join(text_content))
        print(f"Extracted to {out_file} ({os.path.getsize(out_file)} bytes)")
    except Exception as e:
        print(f"Error reading {path}: {e}")
