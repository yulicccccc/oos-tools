# -*- coding: utf-8 -*-
import sys
import os
import fitz
import docx

sys.stdout.reconfigure(encoding='utf-8')
tests_passed = 0
total_tests = 0

def check(condition, desc):
    global tests_passed, total_tests
    total_tests += 1
    if condition:
        tests_passed += 1
        print(f"[PASS] {desc}")
    else:
        print(f"[FAIL] {desc}")

# 1. Master PDF checks
pdf_path = 'CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf'
doc_pdf = fitz.open(pdf_path)
check(len(doc_pdf) == 8, f"Master PDF is strictly 8 pages (got {len(doc_pdf)})")

# Check AcroForm fields
fields = {}
for p in doc_pdf:
    for w in p.widgets():
        fields[w.field_name] = w.field_value

check(fields.get('Text Field12') == 'Kathan Parikh', f"Page 1 Manager Notified is Kathan Parikh (got {fields.get('Text Field12')})")
check('022601-1400' in fields.get('Text Field22', ''), "AMPiScreen Lot 022601-1400 in Text Field22")
check('012503-1142N' in fields.get('Text Field22', ''), "Daily Wash Lot 012503-1142N in Text Field22")
check('022601-1483' in fields.get('Text Field22', ''), "ATP Positive Control Lot 022601-1483 in Text Field22")
check(fields.get('Text Field30') == 'E002222' and fields.get('Text Field31') == 'Dec 2026', "Advance 2 E002222 Cal Due is Dec 2026")
check('ISO 5 BSC E001316' in fields.get('Text Field43', '') and 'ISO 5 BSC E001798' in fields.get('Text Field43', ''), "Primary controls BSC E001316 and E001798 in Text Field43")
check(fields.get('Check Box67') == 'Yes' and fields.get('Check Box69') == 'Off', "EM Conformity: Check Box67=Yes, Check Box69=Off (No ambiguous N/A)")
check(fields.get('Text Field53') == 'Qiyue Chen', "Prepared by is Qiyue Chen")
check(fields.get('Text Field54') == 'Robin Seymour', "Supervisor Approval is Robin Seymour")

# Check Page 5 Narrative
p5_text = fields.get('Text Field51', '')
check('Moraxella osloensis' in p5_text and 'Sporisorium graminicola' in p5_text and 'Staphylococcus haemolyticus' in p5_text, "Page 5 narrative contains confirmed Nanopore species")
check('None of the environmental isolates belong to the genus Bacillus' in p5_text, "Page 5 narrative contains Bacillus genus exclusion defense")

# Check Hyperlinks on Page 7 and 8
p7_links = doc_pdf[6].get_links()
p8_links = doc_pdf[7].get_links()
check(len(p7_links) >= 2, f"Page 7 has active sample hyperlinks (got {len(p7_links)})")
check(len(p8_links) >= 1, f"Page 8 has active EM plate hyperlink (got {len(p8_links)})")
check(any('SPTrZYPqECZmMHXaxXbl9Q__' in l.get('uri', '') for l in p8_links), "Page 8 links to ETX-261021-0578")
doc_pdf.close()

# 2. Desktop Word Table checks
desktop_pcca = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop'
docx_path = os.path.join(desktop_pcca, 'Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx')
doc_word = docx.Document(docx_path)
t3 = doc_word.tables[2]
t3_text = ' '.join([c.text for row in t3.rows for c in row.cells])
check('Date Before Aliquoting' in t3_text, "Table 3 contains Date Before Aliquoting")
check('Date of Aliquoting' in t3_text, "Table 3 contains Date of Aliquoting")
check('Date After Aliquoting' in t3_text, "Table 3 contains Date After Aliquoting")
check('Date of Testing' not in t3_text, "Table 3 does NOT contain Date of Testing (all Aliquoting)")

row13_micro = t3.rows[13].cells[7].text
check('Moraxella osloensis' in row13_micro and 'Sporisorium graminicola' in row13_micro and 'Staphylococcus haemolyticus' in row13_micro and 'Kocuria indica' in row13_micro, "Table 3 Row 13 has confirmed Nanopore species")
check('CFU' not in row13_micro and 'Pending' not in row13_micro, "Table 3 Row 13 has NO CFU counts and NO Pending tag")

# Check Word plate link
r13_links = t3.rows[13].cells[6]._tc.xpath('.//w:hyperlink')
check(len(r13_links) >= 1, f"Word Table 3 Row 13 has plate hyperlink (got {len(r13_links)})")

print(f"\nTotal: {tests_passed}/{total_tests} passed.")
