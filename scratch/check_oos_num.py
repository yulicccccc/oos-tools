import fitz

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261814 GoGoMeds Select (E10747) - ScanRDI - Final Reviewed.pdf"
doc = fitz.open(pdf_path)

for i in range(len(doc)):
    p = doc[i]
    # Check top area: y from 180 to 240, x from 380 to 580
    clip = fitz.Rect(380, 180, 580, 240)
    pix = p.get_pixmap(dpi=150, clip=clip)
    pix.save(f"scratch/top_p{i+1}.png")
    txt = p.get_text("text", clip=clip)
    print(f"Page {i+1} text in top-right clip: {repr(txt)}")
    
    # Also check full page text for 'OOS Number'
    for b in p.get_text("blocks"):
        if "OOS Number" in b[4]:
            print(f"  Page {i+1} block: {b[:4]} -> {repr(b[4])}")
