import docx
from docx.oxml import parse_xml
import pypdf
import fitz

def fix_word_templates():
    for name in ["EM OOS P1 template.docx", "EM OOS P1 template 0.docx"]:
        doc = docx.Document(name)
        t0 = doc.tables[0]
        r5 = t0.rows[5]
        for cell in r5.cells:
            tc_xml = cell._tc.xml
            if '<w:default w:val="1"/>' in tc_xml:
                new_tc_xml = tc_xml.replace('<w:default w:val="1"/>', '<w:default w:val="0"/>')
                parent = cell._tc.getparent()
                new_tc = parse_xml(new_tc_xml)
                parent.replace(cell._tc, new_tc)
                cell._tc = new_tc
        doc.save(name)
        print(f"Fixed checkboxes in Word template: {name}")

def fix_pdf_template():
    pdf_path = "EM OOS P1 template.pdf"
    doc = fitz.open(pdf_path)
    page1 = doc[0]
    for widget in page1.widgets():
        if widget.field_name in ["Check Box0", "Check Box1"]:
            widget.field_value = "Off"
            widget.update()
            print(f"Unchecked {widget.field_name} in {pdf_path}")
    doc.save(pdf_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
    doc.close()
    print("Fixed PDF template checkboxes!")

if __name__ == "__main__":
    fix_word_templates()
    fix_pdf_template()
