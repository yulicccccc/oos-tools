import os
import fitz

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

qyc_path = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")

doc = fitz.open(qyc_path)
page = doc[4]
w = [w for w in page.widgets() if w.field_name == 'Text Field51'][0]

old_text = w.field_value
new_text = old_text.replace(
    "Active air sampling showed recovery of 1 CFU (ETX-260914-0487) in the ISO 8 area, identified as Gram-positive coccobacilli and Gram-positive rods, with zero recovery observed in the ISO 7 buffer or cleanrooms.",
    "Active air sampling showed recovery of 1 CFU (ETX-260914-0487) in the ISO 8 area, identified as Corynebacterium ureicelerivorans (Gram-positive short rods) and Mycobacterium grossiae (Gram-positive rods), with zero recovery observed in the ISO 7 buffer or cleanrooms."
)

assert new_text != old_text, "Replacement failed!"
w.field_value = new_text
w.text_fontsize = 9.2
w.update()

scratch_test_p5 = os.path.join(SCRATCH_DIR, "test_p5_both_micro.pdf")
doc.save(scratch_test_p5)
doc.close()

doc_ver = fitz.open(scratch_test_p5)
p = doc_ver[4]
w_obj = [w for w in p.widgets() if w.field_name == 'Text Field51'][0]
words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
last_y = max([w[3] for w in words]) if words else 0
rem = w_obj.rect.y1 - last_y
last_word = words[-1][4] if words else ''
print(f"Page 5 check with BOTH micro IDs: words={len(words)}, rem={rem:.1f}pt, last_word='{last_word}'")
assert rem > 0, "Overflow on Page 5!"
doc_ver.close()
