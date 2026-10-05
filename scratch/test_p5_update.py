import os
import fitz

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
qyc_path = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")

doc = fitz.open(qyc_path)
page = doc[4]
for w in page.widgets():
    if w.field_name == 'Text Field51':
        old_text = w.field_value
        new_text = old_text.replace(
            "Weekly active air sampling of the cleanroom suite performed on 10Sep2026 showed recovery of 1 CFU (ETX-260921-0520) in the ISO 8 area, identified as Gram-positive cocci, with zero recovery in the ISO 7 buffer or cleanrooms, and weekly surface sampling of the cleanroom suite showed no growth.",
            "Weekly active air sampling of the cleanroom suite performed on 10Sep2026 showed recovery of 1 CFU (ETX-260921-0520) in the ISO 8 area, identified as Micrococcus luteus (Gram-positive cocci), with zero recovery in the ISO 7 buffer or cleanrooms, and weekly surface sampling of the cleanroom suite showed no growth."
        )
        assert new_text != old_text, "Replacement failed!"
        w.field_value = new_text
        w.text_fontsize = 9.2
        w.update()

doc.save(r"scratch\test_p5_micro.pdf")
doc.close()

# Verify bottom margin and overflow
doc_ver = fitz.open(r"scratch\test_p5_micro.pdf")
p = doc_ver[4]
w_obj = [w for w in p.widgets() if w.field_name == 'Text Field51'][0]
words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
last_y = max([w[3] for w in words]) if words else 0
rem = w_obj.rect.y1 - last_y
last_word = words[-1][4] if words else ''
print(f"Page 5 check: words={len(words)}, rem={rem:.1f}pt, last_word='{last_word}'")
assert rem > 0, "Overflow on Page 5!"
doc_ver.close()
