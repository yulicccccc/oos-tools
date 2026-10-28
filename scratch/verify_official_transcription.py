import os, fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
pdf_name = "CORP-FORM-21- 13 Oct 2026 OOS-262329 GoGoMeds Select (E10747) - USP71.pdf"
p_pdf = os.path.join(desktop, pdf_name)

assert os.path.exists(p_pdf), "PDF missing from Desktop!"

doc = fitz.open(p_pdf)
print(f"Total pages: {len(doc)}")
assert len(doc) == 7

# 1. Page 1 checks
p1 = doc[0]
w_tf13 = [w for w in p1.widgets() if w.field_name == 'Text Field13'][0]
print(f"Text Field13 value: '{w_tf13.field_value}'")
assert "Elysse Nioupin, Guanchen Li, Ayomide Odugbesi, and Andrew Carrillo" in w_tf13.field_value, "Full analyst names missing in Text Field13!"
assert "(EN)" not in w_tf13.field_value, "(EN) abbreviation should not be in Text Field13!"

# 2. Page 3 checks
p3 = doc[2]
p3_norm = " ".join(p3.get_text().split())
w_tf48 = [w for w in p3.widgets() if w.field_name == 'Text Field48'][0]
print(f"Text Field48 value: '{w_tf48.field_value}'")
assert w_tf48.field_value == "N/A QYC 28Oct26", "N/A date mismatch!"

assert "to commence incubation." in p3_norm, "Commence incubation missing!"
assert "on 24Sep26 to commence incubation" not in p3_norm, "Old 24Sep26 incubation date still present!"

# 3. Page 4 checks
p4 = doc[3]
p4_norm = " ".join(p4.get_text().split())
assert "(-46/54)" not in p4_norm, "(-46/54) still present!"
assert "(-46/54-A)" not in p4_norm, "(-46/54-A) still present!"
assert "associated with both the processing and subculture activities" not in p4_norm, "Subculture EM sentence still present!"
assert "Additionally, the subculture analyst reported no deviations" not in p4_norm, "Subculture analyst sentence still present!"
assert "verified by laboratory supervisor" in p4_norm, "'verified by laboratory supervisor' missing!"
assert "verified by laboratory management" not in p4_norm, "Old 'verified by laboratory management' still present!"
assert "different from the Gram-positive spore-forming rod" in p4_norm, "'different from' missing!"
assert "distinctly different" not in p4_norm, "'distinctly' still present!"
assert "Monthly cleaning and disinfection of the cleanrooms (ISO 7) and their containing biosafety cabinets (ISO 5) were performed on 13Sep26 for Suite 115" in p4_norm, "Monthly cleaning standard sentence missing!"
assert "approximately 50 samples" not in p4_norm, "Vague 50 samples still present!"
assert "handled another stability sample from the same client" in p4_norm, "Sibling sample missing!"

# 4. Page 5 checks
p5 = doc[4]
p5_norm = " ".join(p5.get_text().split())
assert "Phase II investigation is initiated" not in p5_norm, "Phase II sentence still in Page 5!"
assert "sample retesting is required" not in p5_norm, "Retesting sentence still in Page 5!"
assert "original result is deemed valid" in p5_norm, "Valid conclusion missing!"

# 5. Page 6 checks
p6 = doc[5]
cb87 = [w for w in p6.widgets() if w.field_name == 'Check Box87'][0]
cb88 = [w for w in p6.widgets() if w.field_name == 'Check Box88'][0]
cb89 = [w for w in p6.widgets() if w.field_name == 'Check Box89'][0]
print(f"Check Box87 (Lab Error): {cb87.field_value}")
print(f"Check Box88 (Valid OOS): {cb88.field_value}")
print(f"Check Box89 (Phase II): {cb89.field_value}")
assert cb87.field_value == "Off", "Check Box87 should be Off!"
assert cb88.field_value == "Yes", "Check Box88 should be Yes!"
assert cb89.field_value == "Off", "Check Box89 should be Off!"

# 6. Check Page 7
p7_norm = " ".join(doc[6].get_text().split())
assert "Table 1: Information for ETX-260624-0251 (BUD-T90)" in p7_norm
assert "Table 2: Environmental Monitoring" in p7_norm

print("\n=== ALL AUDIT POINTS & LATEST USER INSTRUCTIONS 100% VERIFIED AND PASSED PERFECTLY! ===")
