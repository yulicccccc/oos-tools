import os
import sys
import docx
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

# Load template / table docx
src_docx = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.docx")
doc = docx.Document(src_docx)

# Table 2: Row 13 (0-indexed)
t2 = doc.tables[1]
r13 = t2.rows[13]

tc5 = r13._tr.xpath('w:tc')[5]  # 1 CFU (ISO 8 114)
tc6 = r13._tr.xpath('w:tc')[6]  # ETX-260914-0487
tc7 = r13._tr.xpath('w:tc')[7]  # Microbial ID

print("tc5 paragraph XML:")
print(tc5.xpath('w:p')[0].xml)

# URL and ID
URL_0487 = "https://etrax.eagleanalytical.com/SubmissionTest/Details/fd3G2StZClcy1TP2ES6BLw__"
SAMPLE_0487 = "ETX-260914-0487"

# Clean tc6 paragraph: replace with perfectly matching paragraph (NO line spacing 360, NO Hyperlink style)
part = doc.part
r_id = part.relate_to(URL_0487, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

# We keep tc6's original <w:tcPr> exactly as is, and replace its <w:p>
for p in tc6.xpath('w:p'):
    tc6.remove(p)

new_p_tc6 = parse_xml(
    f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    f'<w:pPr>'
    f'<w:jc w:val="center"/>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'</w:rPr>'
    f'</w:pPr>'
    f'<w:hyperlink r:id="{r_id}" w:history="1">'
    f'<w:r>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:color w:val="0000FF"/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'<w:u w:val="single"/>'
    f'</w:rPr>'
    f'<w:t>{SAMPLE_0487}</w:t>'
    f'</w:r>'
    f'</w:hyperlink>'
    f'</w:p>'
)
tc6.append(new_p_tc6)

# Now update tc7:
# "不要有&，你就直接两个中间隔开一行就行了，记得要斜体啊，是细菌名"
for p in tc7.xpath('w:p'):
    tc7.remove(p)

new_p_tc7 = parse_xml(
    f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    f'<w:pPr>'
    f'<w:jc w:val="center"/>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:i/>'
    f'<w:iCs/>'
    f'<w:sz w:val="13"/>'
    f'<w:szCs w:val="13"/>'
    f'</w:rPr>'
    f'</w:pPr>'
    f'<w:r>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:i/>'
    f'<w:iCs/>'
    f'<w:sz w:val="13"/>'
    f'<w:szCs w:val="13"/>'
    f'</w:rPr>'
    f'<w:t>Corynebacterium ureicelerivorans</w:t>'
    f'<w:br/>'
    f'<w:br/>'
    f'<w:t>Mycobacterium grossiae</w:t>'
    f'</w:r>'
    f'</w:p>'
)
tc7.append(new_p_tc7)

# Also check Table 3 (Row 13 in Table 3)
t3 = doc.tables[2]
r13_t3 = t3.rows[13]
tc6_t3 = r13_t3._tr.xpath('w:tc')[6]
tc7_t3 = r13_t3._tr.xpath('w:tc')[7]

URL_0520 = "https://etrax.eagleanalytical.com/SubmissionTest/Details/qnwcLQO5BWeJhWCBG7jl8Q__"
SAMPLE_0520 = "ETX-260921-0520"
r_id_0520 = part.relate_to(URL_0520, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

for p in tc6_t3.xpath('w:p'):
    tc6_t3.remove(p)

new_p_tc6_t3 = parse_xml(
    f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    f'<w:pPr>'
    f'<w:jc w:val="center"/>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'</w:rPr>'
    f'</w:pPr>'
    f'<w:hyperlink r:id="{r_id_0520}" w:history="1">'
    f'<w:r>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:color w:val="0000FF"/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'<w:u w:val="single"/>'
    f'</w:rPr>'
    f'<w:t>{SAMPLE_0520}</w:t>'
    f'</w:r>'
    f'</w:hyperlink>'
    f'</w:p>'
)
tc6_t3.append(new_p_tc6_t3)

# Table 3: Micrococcus luteus in italics!
for p in tc7_t3.xpath('w:p'):
    tc7_t3.remove(p)

new_p_tc7_t3 = parse_xml(
    f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    f'<w:pPr>'
    f'<w:jc w:val="center"/>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:i/>'
    f'<w:iCs/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'</w:rPr>'
    f'</w:pPr>'
    f'<w:r>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:i/>'
    f'<w:iCs/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'</w:rPr>'
    f'<w:t>Micrococcus luteus</w:t>'
    f'</w:r>'
    f'</w:p>'
)
tc7_t3.append(new_p_tc7_t3)

out_test_docx = os.path.join(SCRATCH_DIR, "test_bacteria_alignment.docx")
doc.save(out_test_docx)
print("Saved:", out_test_docx)

# Export via Word COM
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

doc_com = word.Documents.Open(out_test_docx)
out_test_pdf = os.path.join(SCRATCH_DIR, "test_bacteria_alignment.pdf")
doc_com.SaveAs2(out_test_pdf, FileFormat=17)
doc_com.Close()
word.Quit()
print("Exported PDF:", out_test_pdf)

# Render cropped image of Table 2 row 13
d = fitz.open(out_test_pdf)
p1 = d[0]
pix = p1.get_pixmap(dpi=150)
pix.save(os.path.join(SCRATCH_DIR, "test_table2_page1.png"))

# Check word coordinates on Table 2 row 13
words = p1.get_text('words')
print("\nWords around ETX-260914-0487:")
for w in words:
    if any(k in w[4] for k in ['0487', 'Corynebacterium', 'grossiae', 'CFU']):
        print(f"  {w[4]:30} y0={w[1]:.2f}, y1={w[3]:.2f}, x0={w[0]:.2f}, x1={w[2]:.2f}")
d.close()
