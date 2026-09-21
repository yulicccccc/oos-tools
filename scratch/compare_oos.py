import fitz
import difflib

pdf_qyc = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261814 GoGoMeds Select (E10747) - ScanRDI - QYC.pdf"
pdf_rs = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261814 GoGoMeds Select (E10747) - ScanRDI - RS Reviewed - Signed by.pdf"

doc_qyc = fitz.open(pdf_qyc)
doc_rs = fitz.open(pdf_rs)

print(f"QYC pages: {len(doc_qyc)}, RS pages: {len(doc_rs)}")

for p_num in range(max(len(doc_qyc), len(doc_rs))):
    text_qyc = doc_qyc[p_num].get_text() if p_num < len(doc_qyc) else ""
    text_rs = doc_rs[p_num].get_text() if p_num < len(doc_rs) else ""
    
    if text_qyc == text_rs:
        print(f"Page {p_num+1}: IDENTICAL")
    else:
        print(f"=== Page {p_num+1}: DIFFERENCES DETECTED ===")
        diff = difflib.unified_diff(
            text_qyc.splitlines(keepends=True),
            text_rs.splitlines(keepends=True),
            fromfile="QYC",
            tofile="RS_Reviewed"
        )
        print("".join(diff))
