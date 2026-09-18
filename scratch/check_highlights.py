import fitz

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261814 GoGoMeds Select (E10747) - ScanRDI - RS Reviewed - Signed by.pdf"
doc = fitz.open(pdf_path)

print("Page count:", len(doc))
page7 = doc[6]

for i, annot in enumerate(page7.annots()):
    rect = annot.rect
    words = page7.get_text("words", clip=rect)
    text = " ".join([w[4] for w in words])
    # Also expand rect slightly by 2 pt to make sure we don't miss slightly off-boundary words
    expanded_rect = fitz.Rect(rect.x0 - 2, rect.y0 - 2, rect.x1 + 2, rect.y1 + 2)
    words_exp = page7.get_text("words", clip=expanded_rect)
    text_exp = " ".join([w[4] for w in words_exp])
    print(f"Highlight {i+1}: Rect={rect}")
    print(f"  Exact text: '{text}'")
    print(f"  Expanded text: '{text_exp}'")
    print(f"  Annot info: {annot.info}")
