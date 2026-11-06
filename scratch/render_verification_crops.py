import fitz
from PIL import Image

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf"
doc = fitz.open(pdf_path)

# 1. Page 1 Text Field3 crop
p1 = doc[0]
pix1 = p1.get_pixmap(dpi=200)
pix1.save("scratch/verify_p1.png")

for field in p1.widgets():
    if field.field_name == 'Text Field3':
        r = field.rect
        scale = 200 / 72
        box = (int(r.x0 * scale) - 20, int(r.y0 * scale) - 20, int(r.x1 * scale) + 20, int(r.y1 * scale) + 20)
        im = Image.open("scratch/verify_p1.png")
        crop = im.crop(box)
        crop.save("scratch/verify_crop_field3.png")
        print("Saved scratch/verify_crop_field3.png")

# 2. Page 7 (Table 1) top-left crop
p7 = doc[6]
pix7 = p7.get_pixmap(dpi=200)
pix7.save("scratch/verify_p7.png")
im7 = Image.open("scratch/verify_p7.png")
# Crop top-left of page 7: x 0 to 300, y 50 to 200
box7 = (int(40 * 200/72), int(60 * 200/72), int(350 * 200/72), int(160 * 200/72))
crop7 = im7.crop(box7)
crop7.save("scratch/verify_crop_tbl1.png")
print("Saved scratch/verify_crop_tbl1.png")

# 3. Page 8 (Table 2) top crop
p8 = doc[7]
pix8 = p8.get_pixmap(dpi=200)
pix8.save("scratch/verify_p8.png")
im8 = Image.open("scratch/verify_p8.png")
box8 = (int(40 * 200/72), int(60 * 200/72), int(550 * 200/72), int(200 * 200/72))
crop8 = im8.crop(box8)
crop8.save("scratch/verify_crop_tbl2_top.png")
print("Saved scratch/verify_crop_tbl2_top.png")

# 4. Page 9 (Table 3) top crop
p9 = doc[8]
pix9 = p9.get_pixmap(dpi=200)
pix9.save("scratch/verify_p9.png")
im9 = Image.open("scratch/verify_p9.png")
box9 = (int(40 * 200/72), int(60 * 200/72), int(550 * 200/72), int(200 * 200/72))
crop9 = im9.crop(box9)
crop9.save("scratch/verify_crop_tbl3_top.png")
print("Saved scratch/verify_crop_tbl3_top.png")
