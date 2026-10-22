import os, sys, shutil
import fitz
import win32com.client as win32
from pypdf import PdfReader, PdfWriter

from build_final_em_tables_v5 import build_perfect_em_table

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

print("=== STEP 1: GENERATE DOCX TABLES ===")
docx_33 = r"scratch\EM table OOS-262133 04SEP2026.docx"
docx_34 = r"scratch\EM table OOS-262134 04SEP2026.docx"
build_perfect_em_table("262133", "262133", docx_33)
build_perfect_em_table("262134", "262134", docx_34)

print("\n=== STEP 2: CONVERT DOCX TO PDF VIA WORD COM ===")
pdf_table_33 = os.path.abspath(r"scratch\EM table OOS-262133 04SEP2026.pdf")
pdf_table_34 = os.path.abspath(r"scratch\EM table OOS-262134 04SEP2026.pdf")

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

for docx_path, pdf_path in [(docx_33, pdf_table_33), (docx_34, pdf_table_34)]:
    wdoc = word.Documents.Open(os.path.abspath(docx_path))
    wdoc.SaveAs2(pdf_path, FileFormat=17)
    wdoc.Close()
    
word.Quit()

# Check page counts
doc33_fitz = fitz.open(pdf_table_33)
doc34_fitz = fitz.open(pdf_table_34)
print(f"Table 262133 PDF pages: {len(doc33_fitz)}")
print(f"Table 262134 PDF pages: {len(doc34_fitz)}")
assert len(doc33_fitz) == 1, f"Table 262133 is {len(doc33_fitz)} pages, must be 1!"
assert len(doc34_fitz) == 1, f"Table 262134 is {len(doc34_fitz)} pages, must be 1!"

print("\n=== STEP 3: ASSEMBLE 7-PAGE MASTER PDF PACKAGES ===")
def assemble_package(oos_id, table_pdf_path, out_pdf_path):
    existing_desktop_pdf = os.path.join(desktop_dir, f"OOS-{oos_id} EM SMO {'143' if oos_id == '262133' else '142'} Air 04SEP2026.pdf")
    
    # Read pages 1-6 from existing desktop PDF (which preserves user's exact edits)
    reader_base = PdfReader(existing_desktop_pdf)
    reader_table = PdfReader(table_pdf_path)
    
    writer = PdfWriter()
    for i in range(6):
        writer.add_page(reader_base.pages[i])
    writer.add_page(reader_table.pages[0])
    
    temp_out = os.path.abspath(f"scratch\\temp_pkg_{oos_id}.pdf")
    with open(temp_out, "wb") as f:
        writer.write(f)
        
    shutil.copyfile(temp_out, out_pdf_path)
    print(f"Assembled and deployed: {out_pdf_path}")
    
    # Verify
    verify_doc = fitz.open(out_pdf_path)
    print(f"  Verified total pages: {len(verify_doc)}")
    assert len(verify_doc) == 7, f"Expected 7 pages, got {len(verify_doc)}"
    
    # Render Page 7 for visual inspection
    pix = verify_doc[6].get_pixmap(dpi=150)
    img_out = f"scratch\\final_verified_p7_{oos_id}.png"
    pix.save(img_out)
    print(f"  Rendered Page 7 verification image: {img_out}")

out_pdf_33 = os.path.join(desktop_dir, "OOS-262133 EM SMO 143 Air 04SEP2026.pdf")
out_pdf_34 = os.path.join(desktop_dir, "OOS-262134 EM SMO 142 Air 04SEP2026.pdf")
assemble_package("262133", pdf_table_33, out_pdf_33)
assemble_package("262134", pdf_table_34, out_pdf_34)

print("\n=== STEP 4: DEPLOY WORD TABLES TO DESKTOP ===")
shutil.copyfile(docx_33, os.path.join(desktop_dir, "EM table OOS-262133 04SEP2026.docx"))
shutil.copyfile(docx_34, os.path.join(desktop_dir, "EM table OOS-262134 04SEP2026.docx"))

print("\nAll deliverables assembled and deployed successfully!")
