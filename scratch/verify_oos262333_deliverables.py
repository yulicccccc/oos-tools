import fitz
import os
import docx

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"

DOCX_NAME = "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx"
PDF_NAME = "CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf"

docx_desktop = os.path.join(DESKTOP_DIR, DOCX_NAME)
docx_docs = os.path.join(DOCS_DIR, DOCX_NAME)
pdf_desktop = os.path.join(DESKTOP_DIR, PDF_NAME)
pdf_docs = os.path.join(DOCS_DIR, PDF_NAME)

print("=== [TEST 1] File Existence & Synchronization ===")
assert os.path.exists(docx_desktop), f"Missing {docx_desktop}"
assert os.path.exists(docx_docs), f"Missing {docx_docs}"
assert os.path.exists(pdf_desktop), f"Missing {pdf_desktop}"
assert os.path.exists(pdf_docs), f"Missing {pdf_docs}"
print("PASS: Both DOCX and PDF deliverables exist in Desktop and Documents.")

print("\n=== [TEST 2] DOCX Structure Verification ===")
doc = docx.Document(docx_desktop)
assert len(doc.tables) == 3, f"Expected 3 tables, got {len(doc.tables)}"
t1, t2, t3 = doc.tables
assert len(t1.rows) == 2, f"Table 1 expected 2 rows, got {len(t1.rows)}"
assert len(t1.columns) == 6, f"Table 1 expected 6 columns, got {len(t1.columns)}"
assert len(t2.rows) == 16, f"Table 2 expected 16 rows, got {len(t2.rows)}"
assert len(t2.columns) == 9, f"Table 2 expected 9 columns, got {len(t2.columns)}"
assert len(t3.rows) == 16, f"Table 3 expected 16 rows, got {len(t3.rows)}"
assert len(t3.columns) == 9, f"Table 3 expected 9 columns, got {len(t3.columns)}"

# Check Table 1
assert t1.rows[1].cells[0].text.strip() == "Alex Saravia"
assert t1.rows[1].cells[1].text.strip() == "Cuong Du"
assert "ETX-261002-0288" in t1.rows[1].cells[2].text
assert "ETX-261012-0654" in t1.rows[1].cells[3].text
assert "1 x 300mL FTM" in t1.rows[1].cells[4].text
assert "Bacillus paralicheniformis" in t1.rows[1].cells[5].text

# Check Table 3 Row 13 active air
row13_t3 = [c.text.strip() for c in t3.rows[13].cells]
assert "6 CFUs" in row13_t3[5]
assert "ETX-261021-0578" in row13_t3[6]
assert "Moraxella osloensis" in row13_t3[7]
assert "Kocuria rhizophila" in row13_t3[7]
print("PASS: DOCX content and table structure verified.")

print("\n=== [TEST 3] Master PDF Verification ===")
p = fitz.open(pdf_desktop)
print(f"Master PDF page count: {len(p)} pages.")
assert len(p) == 8, f"Expected strictly 8 pages master PDF, got {len(p)}"

# Verify Page 1 AcroForm fields
p1 = p[0]
widgets = {w.field_name: w for w in p1.widgets()}
assert "Mukyung Jang" in widgets["Text Field3"].field_value
assert "Alex Saravia" in widgets["Text Field3"].field_value
assert "Cuong Du" in widgets["Text Field3"].field_value
assert widgets["Text Field12"].field_value == "Kathan Parikh"
assert widgets["Date Field0"].field_value == "05-Oct-2026"
assert widgets["Date Field1"].field_value == "12-Oct-2026"
assert widgets["Date Field2"].field_value == "12-Oct-2026"
assert widgets["Date Field3"].field_value == "12-Oct-2026"

# Verify Check Box 84 is Off
cb84 = None
for page in p[:6]:
    for w in page.widgets():
        if w.field_name == "Check Box84":
            cb84 = w.field_value
assert cb84 == "Off", f"Check Box84 must be strictly Off, got {cb84}"
print("PASS: Check Box84 is strictly Off.")

# Verify Page 6 Sign-Off
p6 = p[5]
w6 = {w.field_name: w for w in p6.widgets()}
assert w6["Text Field53"].field_value == "Qiyue Chen"
assert w6["Text Field54"].field_value == "Robin Seymour"
print("PASS: Prepared by Qiyue Chen and Supervisor Approval Robin Seymour verified.")

# Verify Page 7 Table 1 & Table 2 layout
p7 = p[6]
blocks_p7 = p7.get_text("blocks")
t1_title = [b for b in blocks_p7 if "Table 1:" in b[4]]
assert len(t1_title) > 0, "Table 1 title not found on Page 7"
t2_title = [b for b in blocks_p7 if "Table 2:" in b[4]]
assert len(t2_title) > 0, "Table 2 title not found on Page 7"

# Check hyperlink on Page 7
links_p7 = p7.get_links()
assert any("ETX-261002-0288" in l.get("uri", "") or "Details" in l.get("uri", "") for l in links_p7), "Missing hyperlink on Page 7"
print(f"PASS: Page 7 Table 1 & 2 layout verified with {len(links_p7)} link(s).")

# Verify Page 8 Table 3 layout
p8 = p[7]
blocks_p8 = p8.get_text("blocks")
t3_title = [b for b in blocks_p8 if "Table 3:" in b[4]]
assert len(t3_title) > 0, "Table 3 title not found on Page 8"

# Check hyperlink on Page 8
links_p8 = p8.get_links()
assert any("ETX-261021-0578" in l.get("uri", "") or "Details" in l.get("uri", "") for l in links_p8), "Missing hyperlink on Page 8"
print(f"PASS: Page 8 Table 3 verified with {len(links_p8)} link(s).")

print("\n==============================================================")
print("ALL AUTOMATED VERIFICATION TESTS PASSED (100% GREEN)!")
print("==============================================================")
