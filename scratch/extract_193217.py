import fitz, glob, os
from PIL import Image

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\svc-scan@eagleanalytical.com_20260925_193217.pdf")
os.makedirs("scratch/scans/193217", exist_ok=True)

for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=120)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img.save(f"scratch/scans/193217/p{i+1}.png")
    # Save header crop
    hdr = img.crop((0, 0, pix.width, min(220, pix.height)))
    hdr.save(f"scratch/scans/193217/p{i+1}_hdr.png")

print(f"Extracted all {len(doc)} pages and headers to scratch/scans/193217/")
