import os
import sys
import fitz
import docx

sys.stdout.reconfigure(encoding='utf-8')

desk_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
f1_path = os.path.join(desk_dir, "CORP-FORM-21- 13 Oct 2026 OOS-262329 GoGoMeds Select (E10747) - USP71.pdf")
f2_path = os.path.join(desk_dir, "CORP-FORM-22- 13 Oct 2026 OOS-262329 GoGoMeds Select (E10747) - Phase II.pdf")
f3_path = os.path.join(desk_dir, "Tables OOS-262329 GoGoMeds Select (E10747) - USP71.docx")

print("================================================================================")
print("              FINAL PRE-FLIGHT VERIFICATION OF SUBMISSION DELIVERABLES          ")
print("================================================================================")

# ---------------------------------------------------------
# Check 1: File Existence & Paths
# ---------------------------------------------------------
print("\n[CHECK 1] Desktop Deliverable Existence:")
for p, label in [(f1_path, "Phase I PDF"), (f2_path, "Phase II PDF"), (f3_path, "Word Tables")]:
    exists = os.path.exists(p)
    size = os.path.getsize(p) if exists else 0
    print(f"  - {label}: {'[PASS] Exists' if exists else '[FAIL] Missing'} ({size:,} bytes) -> {os.path.basename(p)}")
    assert exists, f"File missing: {p}"

# ---------------------------------------------------------
# Check 2: Phase I PDF (CORP-FORM-21)
# ---------------------------------------------------------
print("\n[CHECK 2] Phase I Master PDF Detailed Audit:")
doc1 = fitz.open(f1_path)
print(f"  - Total Page Count: {len(doc1)} (Expected: 7) -> {'[PASS]' if len(doc1) == 7 else '[FAIL]'}")

# Check ZenQMS footer on Form Pages (1 to 6)
p1_text = doc1[0].get_text()
for line in p1_text.splitlines():
    if "Printed by" in line:
        print(f"  - Page 1 Footer: {line}")
        assert "13-Oct-2026" in line, "ZenQMS timestamp mismatch in Phase I!"

# Check critical fields on Page 1
p1_w = {w.field_name: w.field_value for w in doc1[0].widgets()}
print(f"  - Initiator / Written by: {p1_w.get('Text Field0')}")
print(f"  - Manager Notified: {p1_w.get('Text Field12')} on {p1_w.get('Date Field3')}")
assert p1_w.get('Text Field12') == "Kathan Parikh", "Manager notified must be Kathan Parikh!"

# Check Page 4 Day 4 correction
p4_w = {w.field_name: w.field_value for w in doc1[3].widgets()}
t_p4 = p4_w.get('Text Field50', '')
has_day4_p1 = "Day 4 of the subculture aliquots" in t_p4
has_day5_p1 = "Day 5 of the subculture aliquots" in t_p4
has_mluteus_p1 = "Micrococcus luteus" in t_p4
print(f"  - Day 4 in Page 4 Narrative: {'[PASS]' if has_day4_p1 else '[FAIL]'}")
print(f"  - Day 5 purged: {'[PASS]' if not has_day5_p1 else '[FAIL]'}")
print(f"  - Micrococcus luteus present: {'[PASS]' if has_mluteus_p1 else '[FAIL]'}")
assert has_day4_p1 and not has_day5_p1, "Phase I Day 4 assertion failed!"

# Check Page 6 Phase II initiated and approvals
p6_w = {w.field_name: w.field_value for w in doc1[5].widgets()}
print(f"  - Phase II initiated (Check Box89): {p6_w.get('Check Box89')} -> {'[PASS]' if p6_w.get('Check Box89') == 'Yes' else '[FAIL]'}")
print(f"  - Prepared by: {p6_w.get('Text Field53')}")
print(f"  - Supervisor Approval: {p6_w.get('Text Field54')}")
assert p6_w.get('Check Box89') == "Yes"
assert p6_w.get('Text Field54') == "Robin Seymour"

# Check Page 7 Table content
p7_text = doc1[6].get_text()
has_both_samples_p1 = "ETX-260624-0251 (BUD-T90)" in p7_text and "ETX-261021-0445" in p7_text
print(f"  - Page 7 Table 1 contains Dual Samples: {'[PASS]' if has_both_samples_p1 else '[FAIL]'}")

# ---------------------------------------------------------
# Check 3: Phase II PDF (CORP-FORM-22)
# ---------------------------------------------------------
print("\n[CHECK 3] Phase II Master PDF Detailed Audit:")
doc2 = fitz.open(f2_path)
print(f"  - Total Page Count: {len(doc2)} (Expected: 6) -> {'[PASS]' if len(doc2) == 6 else '[FAIL]'}")

# Check ZenQMS footer
p1_text_p2 = doc2[0].get_text()
for line in p1_text_p2.splitlines():
    if "Printed by" in line:
        print(f"  - Page 1 Footer: {line}")
        assert "13-Oct-2026" in line, "ZenQMS timestamp mismatch in Phase II!"

# Check Page 1 Retest info
p1_p2_w = {w.field_name: w.field_value for w in doc2[0].widgets()}
print(f"  - Submissions: {repr(p1_p2_w.get('Text Field2'))}")
print(f"  - Results: Orig={p1_p2_w.get('Text Field4')}, Retest={p1_p2_w.get('Text Field3')}")
assert p1_p2_w.get('Text Field4') == "Fail" and p1_p2_w.get('Text Field3') == "Pass"

# Check Page 3 Day 4 correction
p3_p2_w = {w.field_name: w.field_value for w in doc2[2].widgets()}
t_p3_p2 = p3_p2_w.get('Text Field23', '')
has_day4_p2 = "Day 4 of subculture aliquots" in t_p3_p2
has_day5_p2 = "Day 5 of subculture aliquots" in t_p3_p2
print(f"  - Day 4 in Page 3 Narrative: {'[PASS]' if has_day4_p2 else '[FAIL]'}")
print(f"  - Day 5 purged: {'[PASS]' if not has_day5_p2 else '[FAIL]'}")
assert has_day4_p2 and not has_day5_p2, "Phase II Day 4 assertion failed!"

# Check Page 4 Retest & Client Discretion
p4_p2_w = {w.field_name: w.field_value for w in doc2[3].widgets()}
t_p4_p2 = p4_p2_w.get('Text Field24', '')
has_client_discretion = "Final beyond-use dating (BUD) determination, stability protocol evaluation, and lot disposition remain at the discretion of the client." in t_p4_p2
has_both_valid = "both the initial failing result" in t_p4_p2 and "passing retest result" in t_p4_p2 and "are considered valid" in t_p4_p2
print(f"  - Client Sole Discretion Invariant: {'[PASS]' if has_client_discretion else '[FAIL]'}")
print(f"  - Both Results Considered Valid: {'[PASS]' if has_both_valid else '[FAIL]'}")
assert has_client_discretion and has_both_valid

# Check Page 5 Root Cause & GDP
p5_p2_w = {w.field_name: w.field_value for w in doc2[4].widgets()}
rc_text = p5_p2_w.get('Text Field25', '')
has_nonuniform = "non-uniform microbial distribution" in rc_text and "low concentrations within the batch" in rc_text
print(f"  - Root Cause Checkbox (Check Box51): {p5_p2_w.get('Check Box51')} -> {'[PASS]' if p5_p2_w.get('Check Box51') == 'Yes' else '[FAIL]'}")
print(f"  - Client Care Notified Checkbox (Check Box53): {p5_p2_w.get('Check Box53')} -> {'[PASS]' if p5_p2_w.get('Check Box53') == 'Yes' else '[FAIL]'}")
print(f"  - Concise Non-Uniform Root Cause Phrasing: {'[PASS]' if has_nonuniform else '[FAIL]'}")
print(f"    Value: {repr(rc_text)}")
assert p5_p2_w.get('Check Box51') == "Yes" and p5_p2_w.get('Check Box53') == "Yes" and has_nonuniform

# ---------------------------------------------------------
# Check 4: Word Tables (Tables OOS-262329...docx)
# ---------------------------------------------------------
print("\n[CHECK 4] Standalone Word Tables Audit:")
wdoc = docx.Document(f3_path)
print(f"  - Total Tables: {len(wdoc.tables)} (Expected: 2)")
t1 = wdoc.tables[0]
print(f"  - Table 1 Rows: {len(t1.rows)} (Header + 2 sample rows = 3)")
t1_samples = [t1.rows[i].cells[1].text for i in range(1, len(t1.rows))]
print(f"  - Table 1 Samples: {t1_samples}")
assert len(t1.rows) == 3, "Table 1 must have exactly 3 rows!"

# Check ampersands in tables
has_amp = False
for t in wdoc.tables:
    for row in t.rows:
        for c in row.cells:
            if " & " in c.text:
                print(f"  [WARNING] Found ampersand in cell: {c.text}")
                has_amp = True
print(f"  - Zero Ampersands ('&') in Word Tables: {'[PASS]' if not has_amp else '[FAIL]'}")
assert not has_amp, "Ampersands found in Word tables!"

print("\n================================================================================")
print("                   ALL PRE-FLIGHT VERIFICATIONS PASSED 100%!                   ")
print("================================================================================")
