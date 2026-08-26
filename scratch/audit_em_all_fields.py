import pypdf
import glob
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

sample_paths = [
    r"G:\CRO\OOS\2026\05-2026\OOS- 260422 ScanCO HS BSC1309 S3 17FEB2026 and ScanCO HS BSC1309 S1 17FEB2026 MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260469 Scan GA 1311 SettR 23FEB2026-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260541 Scan PG BSC1310 Sett2 26FEB2026-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260542 Scan DH BSC1313 SettL 26FEB2026 draft-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260543 Scan VV BSC1309 S1 27FEB2026  draft-MN-KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260361 Sterility GS BSC1314 Sett2 05FEB2026-signed by MN - KSP.pdf",
    r"G:\CRO\OOS\2026\05-2026\OOS- 260365 EM OA 114A cart 10FEB2026-Signed by MN - KSP.pdf"
]

print("==========================================================")
print("AUDITING REAL SAMPLE PDF FORM FIELDS (FIRST 60 FIELDS)")
print("==========================================================")

field_audit_summary = {}

for idx, ppath in enumerate(sample_paths):
    if not os.path.exists(ppath): continue
    reader = pypdf.PdfReader(ppath)
    fields = reader.get_fields()
    print(f"\n--- SAMPLE {idx+1}: {os.path.basename(ppath)} ---")
    for key, fobj in fields.items():
        val = fobj.get('/V', '')
        if val:
            # Store in audit dictionary
            if key not in field_audit_summary:
                field_audit_summary[key] = []
            field_audit_summary[key].append(str(val).strip())

print("\n\n==========================================================")
print("FIELD-BY-FIELD AUDIT DICTIONARY ACROSS ALL 7 SAMPLES:")
print("==========================================================")
for key, vals in field_audit_summary.items():
    print(f"Key: '{key}' ({len(vals)} occurrences)")
    for v in vals[:2]: # Show first 2 examples
        snippet = repr(v)[:120]
        print(f"   -> Example: {snippet}")
