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
for p in doc_pdf[:6]:
    for w in p.widgets():
        fields[w.field_name] = w.field_value

check(fields.get('Text Field57') == '262333', "OOS ID is 262333")
check(fields.get('Date Field0') == '05-Oct-2026', "Page 1 Testing Date is 05-Oct-2026")
check(fields.get('Date Field1') == '12-Oct-2026', "Page 1 Read Date is 12-Oct-2026")
check(fields.get('Text Field2') == 'ETX-261002-0288', "Page 1 Sample ID is ETX-261002-0288")
check(fields.get('Text Field4') == 'TESAMORELIN 8 MG/ML (3 ML) Injection', "Page 1 Sample Description is TESAMORELIN 8 MG/ML (3 ML) Injection")
check(fields.get('Text Field5') == 'Injection', "Page 1 Dosage Form is Injection")
check(fields.get('Text Field6') == 'LG342011076', "Page 1 Lot Number is LG342011076")
check(fields.get('Text Field12') == 'Kathan Parikh', f"Page 1 Manager Notified is Kathan Parikh (got {fields.get('Text Field12')})")
check(fields.get('Check Box34') == 'Yes' and fields.get('Check Box36') == 'Off', "Page 1 Storage: Check Box34=Yes, Check Box36=Off")
check('022601-1400' in fields.get('Text Field22', ''), "AMPiScreen Lot 022601-1400 in Text Field22")
check('012503-1142N' in fields.get('Text Field22', ''), "Daily Wash Lot 012503-1142N in Text Field22")
check('022601-1483' in fields.get('Text Field22', ''), "ATP Positive Control Lot 022601-1483 in Text Field22")
check(fields.get('Text Field25') == '022601-1483', "Page 2 ATP Control Lot in Text Field25 is 022601-1483")
check(fields.get('Text Field30') == 'E002222' and fields.get('Text Field31') == 'Dec 2026', "Advance 2 E002222 Cal Due is Dec 2026")
check('E001736 (CR 114)' in fields.get('Text Field32', ''), "Suite 114 equipment ID E001736 in Text Field32")
check('ISO 5 BSC E001316' in fields.get('Text Field43', '') and 'ISO 5 BSC E001798' in fields.get('Text Field43', ''), "Primary controls BSC E001316 and E001798 in Text Field43")
check(fields.get('Check Box67') == 'Yes' and fields.get('Check Box69') == 'Off', "EM Conformity: Check Box67=Yes, Check Box69=Off (No ambiguous N/A)")
check(fields.get('Check Box84') == 'Off', "Page 3 Check Box84 is strictly Off")
check(fields.get('Text Field48') == 'N/A QYC 22Oct26', "Page 3 Text Field48 has GDP sign-off N/A QYC 22Oct26")
check(fields.get('Text Field53') == 'Qiyue Chen', "Prepared by is Qiyue Chen")
check(fields.get('Text Field54') == 'Robin Seymour', "Supervisor Approval is Robin Seymour")

# Check narrative content
t49 = fields.get('Text Field49', '')
t50 = fields.get('Text Field50', '')
t51 = fields.get('Text Field51', '')
check(len(t49) > 2500, f"Page 3 narrative is full ({len(t49)} chars)")
check(len(t50) > 2500, f"Page 4 narrative is full ({len(t50)} chars)")
check(len(t51) > 2500, f"Page 5 narrative is full ({len(t51)} chars)")
check('Moraxella osloensis' in t51 and 'Ustilago maydis' in t51, "Page 5 narrative contains confirmed Nanopore species")
check('genus Bacillus' in t51, "Page 5 narrative contains Bacillus genus exclusion defense")

# Check Tables on Page 7 & 8
p7_text = doc_pdf[6].get_text()
p8_text = doc_pdf[7].get_text()
check('Aliquoting Analyst' in p7_text, "Table 1 Column 2 is Aliquoting Analyst")
check('Bacillus paralicheniformis' in p7_text and 'Gram-positive rod' not in p7_text, "Table 1 Microbial ID has pure species without Gram stain")
check('Date Before' in p8_text and 'Aliquoting' in p8_text, "Table 3 contains Date Before Aliquoting")
check('Date of' in p8_text and 'Aliquoting' in p8_text, "Table 3 contains Date of Aliquoting")
check('Date After' in p8_text and 'Aliquoting' in p8_text, "Table 3 contains Date After Aliquoting")
check('Date of Testing' not in p8_text, "Table 3 does NOT contain Date of Testing")
check('Moraxella osloensis' in p8_text and 'Ustilago maydis' in p8_text, "Table 3 has confirmed Nanopore species")

# Check hyperlinks in PDF
links_p7 = doc_pdf[6].get_links()
links_p8 = doc_pdf[7].get_links()
check(len(links_p7) >= 1 and any('ETX-261002-0288' in str(l) or 'Q57eK3N9WcR0Qh' in str(l) for l in links_p7), "Page 7 has active Sample ID hyperlink")
check(len(links_p8) >= 1 and any('SPTrZYPqECZmMHXaxXbl9Q' in str(l) or 'ETX-261021-0578' in str(l) for l in links_p8), "Page 8 has active Active Air plate hyperlink")
check(len(doc_pdf[6].get_drawings()) > 500, f"Page 7 has table grid borders (drawings count: {len(doc_pdf[6].get_drawings())})")
check(len(doc_pdf[7].get_drawings()) > 500, f"Page 8 has table grid borders (drawings count: {len(doc_pdf[7].get_drawings())})")
doc_pdf.close()

# 2. Desktop Word Table checks
desktop_pcca = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop'
docx_path = os.path.join(desktop_pcca, 'Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx')
doc_word = docx.Document(docx_path)
check(len(doc_word.tables) == 3, f"Word document has strictly 3 tables (got {len(doc_word.tables)})")

for idx, t in enumerate(doc_word.tables):
    check(t._tbl.tblPr.find(docx.oxml.ns.qn('w:tblBorders')) is not None, f"Word Table {idx+1} has tblBorders defined")

t1 = doc_word.tables[0]
check(t1.rows[0].cells[1].text.strip() == 'Aliquoting Analyst', "Word Table 1 Column 2 is Aliquoting Analyst")
check(t1.rows[1].cells[5].text.strip() == 'Bacillus paralicheniformis', "Word Table 1 Microbial ID is pure species name")

t3 = doc_word.tables[2]
t3_text = ' '.join([c.text for row in t3.rows for c in row.cells])
check('Date Before Aliquoting' in t3_text, "Word Table 3 contains Date Before Aliquoting")
check('Date of Aliquoting' in t3_text, "Word Table 3 contains Date of Aliquoting")
check('Date After Aliquoting' in t3_text, "Word Table 3 contains Date After Aliquoting")
check('Date of Testing' not in t3_text, "Word Table 3 does NOT contain Date of Testing")

row13_micro = t3.rows[13].cells[7].text
check('Moraxella osloensis' in row13_micro and 'Ustilago maydis' in row13_micro, "Word Table 3 Row 13 has confirmed Nanopore species")
check('CFU' not in row13_micro and 'Pending' not in row13_micro, "Word Table 3 Row 13 has NO CFU counts and NO Pending tag")

print(f"\nTotal: {tests_passed}/{total_tests} passed.")
assert tests_passed == total_tests, f"Some tests failed! ({tests_passed}/{total_tests})"
