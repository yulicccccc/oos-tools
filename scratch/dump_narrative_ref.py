import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf")
print("=== PAGE 3 TEXT ===")
print(doc[2].get_text("text"))
print("=== PAGE 4 TEXT ===")
print(doc[3].get_text("text"))
print("=== PAGE 5 TEXT ===")
print(doc[4].get_text("text"))
print("=== PAGE 6 TEXT ===")
print(doc[5].get_text("text"))
