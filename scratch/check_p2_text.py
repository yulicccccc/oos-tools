import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf")
p2 = doc[1]
# Text Field43 rect is roughly x: 396 to 493, y: 587 to 737
# Text Field44 rect is roughly x: 494 to 601, y: 587 to 737
print("=== TEXT IN RECT 43 ===")
print(p2.get_text("text", clip=fitz.Rect(390, 580, 495, 740)))
print("=== TEXT IN RECT 44 ===")
print(p2.get_text("text", clip=fitz.Rect(494, 580, 605, 740)))
