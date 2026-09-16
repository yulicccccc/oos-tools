import pypdf, difflib

qyc_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261814 GoGoMeds Select (E10747) - ScanRDI - QYC.pdf"
our_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS-261814 GoGoMeds Select (E10747) - ScanRDI.pdf"

qyc_reader = pypdf.PdfReader(qyc_path)
our_reader = pypdf.PdfReader(our_path)

print(f"QYC pages: {len(qyc_reader.pages)}, Our pages: {len(our_reader.pages)}")


# 1. Field differences
qyc_fields = qyc_reader.get_fields() or {}
our_fields = our_reader.get_fields() or {}

print("\n--- FIELD COMPARISON ---")
diff_fields = []
for k in set(qyc_fields.keys()) | set(our_fields.keys()):
    v_qyc = qyc_fields.get(k, {}).get('/V')
    v_our = our_fields.get(k, {}).get('/V')
    if v_qyc != v_our:
        diff_fields.append((k, v_our, v_qyc))

print(f"Total field differences: {len(diff_fields)}")
for k, v_our, v_qyc in diff_fields:
    print(f"\n[Field: {k}]")
    print(f"  OUR: {repr(v_our)}")
    print(f"  QYC: {repr(v_qyc)}")

# 2. Text extraction diff
print("\n--- PAGE TEXT DIFFERENCES ---")
for i in range(max(len(qyc_reader.pages), len(our_reader.pages))):
    t_qyc = qyc_reader.pages[i].extract_text() if i < len(qyc_reader.pages) else ""
    t_our = our_reader.pages[i].extract_text() if i < len(our_reader.pages) else ""
    if t_qyc.strip() != t_our.strip():
        print(f"\n=== Page {i+1} Text Diff ===")
        diff = list(difflib.ndiff(t_our.splitlines(), t_qyc.splitlines()))
        for line in diff:
            if line.startswith(('+', '-')):
                print(line)


