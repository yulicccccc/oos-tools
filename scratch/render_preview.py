import fitz

doc = fitz.open(r'scratch\OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf')
for pno in [2, 3, 4]:
    page = doc[pno]
    pix = page.get_pixmap(dpi=150)
    pix.save(f'scratch/check_p{pno+1}_human.png')
print("Rendered check_p3_human.png, check_p4_human.png, check_p5_human.png")
