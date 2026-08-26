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

scratch_dir = r"C:\Users\qchen\.gemini\antigravity\brain\f2811f7a-6e22-4e06-be5b-f4e86a5e371a\scratch"

for i, path in enumerate(pdf_paths):
    out_file = os.path.join(scratch_dir, f"pdf_{i+1}_fields.txt")
    print(f"Processing PDF {i+1}: {os.path.basename(path)} -> {out_file}")
    reader = pypdf.PdfReader(path)
    fields = reader.get_fields()
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(f"PDF FILE: {os.path.basename(path)}\n")
        if fields:
            f.write(f"Total AcroForm Fields: {len(fields)}\n\n")
            for field_name, field_val in fields.items():
                val = field_val.get('/V')
                if val:
                    f.write(f"Field: '{field_name}'\nValue:\n{val}\n{'-'*40}\n")
        else:
            f.write("No AcroForm fields found.\n")
