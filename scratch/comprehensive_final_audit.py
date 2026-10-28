import os, sys, fitz, docx

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
pdf_name = "CORP-FORM-21- 13 Oct 2026 OOS-262329 GoGoMeds Select (E10747) - USP71.pdf"
docx_name = "Tables OOS-262329 GoGoMeds Select (E10747) - USP71.docx"

pdf_path = os.path.join(desktop, pdf_name)
docx_path = os.path.join(desktop, docx_name)

print("="*60)
print("COMPREHENSIVE FINAL AUDIT FOR OOS-262329 SUBMISSION")
print("="*60)

# Check 1: Existence and File Integrity
assert os.path.exists(pdf_path), "PDF file missing from Desktop!"
assert os.path.exists(docx_path), "Word file missing from Desktop!"
print("[PASS 1] Both deliverables exist on Canonical Desktop.")

doc = fitz.open(pdf_path)
assert len(doc) == 7, f"Expected 7 pages, got {len(doc)}"
print("[PASS 2] PDF page count strictly equals 7.")

# Check 2: Page 1 Fields
p1 = doc[0]
w_p1 = {w.field_name: w for w in p1.widgets()}

assert w_p1['Text Field57'].field_value == '262329', "OOS ID mismatch!"
assert w_p1['Date Field0'].field_value == '23-Sep-2026', "Processing date mismatch!"
assert w_p1['Date Field1'].field_value == '13-Oct-2026', "Incident/Test date mismatch!"
assert w_p1['Date Field2'].field_value == '13-Oct-2026', "Incident/Test date mismatch!"
assert w_p1['Date Field3'].field_value == '13-Oct-2026', "Incident/Test date mismatch!"

assert w_p1['Text Field0'].field_value == 'Guanchen Li (Written by Qiyue Chen)', "Text Field0 mismatch!"
assert w_p1['Text Field1'].field_value == 'USP <71> / EP 2.6.1 Sterility Test', "Test name mismatch!"
assert w_p1['Text Field2'].field_value == 'ETX-260624-0251 (BUD-T90)', "Sample ID mismatch!"
assert w_p1['Text Field4'].field_value == 'Sermorelin 0.9mg/ml', "Analyte mismatch!"
assert w_p1['Text Field5'].field_value == 'Liquid', "Dosage form mismatch!"
assert w_p1['Text Field6'].field_value == '17009', "Lot number mismatch!"
assert w_p1['Text Field8'].field_value == 'MICRO-SOP-5', "SOP mismatch!"
assert w_p1['Text Field9'].field_value == '24Jul2026', "SOP effective date mismatch!"
assert w_p1['Text Field10'].field_value == '19', "SOP revision mismatch!"
assert w_p1['Text Field11'].field_value == 'Pass/Fail', "Specification mismatch!"
assert w_p1['Text Field12'].field_value == 'Robin Seymour', "Supervisor mismatch!"

# Text Field13 Analyst Names
tf13 = w_p1['Text Field13'].field_value
assert tf13 == "Yes, analysts Elysse Nioupin, Guanchen Li, Ayomide Odugbesi, and Andrew Carrillo were comprehensively interviewed.", f"Text Field13 mismatch: '{tf13}'"
assert "(EN)" not in tf13 and "(GL)" not in tf13 and "(AOD)" not in tf13 and "(ANC)" not in tf13, "Informal initials found in Text Field13!"
print("[PASS 3] Page 1 fields, dates, SOP, supervisor, and full analyst names verified.")

# Check 3: Page 2 Equipment & Cleanrooms
p2 = doc[1]
w_p2 = {w.field_name: w for w in p2.widgets()}
assert w_p2['Text Field30'].field_value == 'ISO 5 BSC E001314', "BSC mismatch!"
assert w_p2['Text Field31'].field_value == 'Jun 2027', "BSC calibration mismatch!"
assert w_p2['Text Field32'].field_value == 'E001737 (CR115)', "Suite 115 cleanroom ID mismatch!"
assert w_p2['Text Field33'].field_value == 'Dec 2026', "Cleanroom certification date mismatch!"
assert "Incubator E001356" in w_p2['Text Field43'].field_value, "Incubator E001356 missing!"
assert "Incubator E001357" in w_p2['Text Field43'].field_value, "Incubator E001357 missing!"
assert "BSC" not in w_p2['Text Field43'].field_value, "BSC should NOT be duplicated in general equipment Text Field43!"
assert "Jan 2027" in w_p2['Text Field44'].field_value, "Incubator calibration mismatch!"
assert "Jun 2027" not in w_p2['Text Field44'].field_value, "BSC calibration Jun 2027 should NOT be in Text Field44!"
print("[PASS 4] Page 2 equipment, calibrations, and cleanroom IDs verified.")

# Check 4: Page 3 Narrative & GDP Sign-off
p3 = doc[2]
w_p3 = {w.field_name: w for w in p3.widgets()}
tf48 = w_p3['Text Field48'].field_value
assert tf48 == "N/A QYC 28Oct26", f"Text Field48 GDP sign-off mismatch: '{tf48}'"

p3_text = " ".join(p3.get_text().split())
assert "to commence incubation." in p3_text, "commence incubation missing!"
assert "on 24Sep26 to commence incubation" not in p3_text, "Old 24Sep26 incubation date still present!"
assert "midnight" not in p3_text, "Unwanted midnight text present!"
assert "ETX-260401-0599" in p3_text, "Method suitability ETX number missing!"
assert "Elysse Nioupin" in p3_text and "Guanchen Li" in p3_text, "Prep and process analysts missing in Page 3!"
print("[PASS 5] Page 3 narrative and GDP sign-off verified.")

# Check 5: Page 4 Narrative Audit
p4 = doc[3]
p4_text = " ".join(p4.get_text().split())

# Prohibited items check
assert "(-46/54)" not in p4_text, "Internal bottle index (-46/54) found!"
assert "(-46/54-A)" not in p4_text, "Internal bottle index (-46/54-A) found!"
assert "associated with both the processing and subculture activities" not in p4_text, "Unverified subculture EM plate sentence found!"
assert "Additionally, the subculture analyst reported no deviations" not in p4_text, "Subculture analyst deviation report sentence found!"
assert "distinctly" not in p4_text, "Subjective word 'distinctly' found!"
assert "approximately 50 samples" not in p4_text, "Vague 50 samples found!"

# Required items check
assert "verified by laboratory supervisor" in p4_text, "'verified by laboratory supervisor' missing!"
assert "different from the Gram-positive spore-forming rod" in p4_text, "'different from' morphology phrasing missing!"
assert "Bacillus megaterium" in p4_text, "Microbial identity missing!"
assert "Monthly cleaning and disinfection of the cleanrooms (ISO 7) and their containing biosafety cabinets (ISO 5) were performed on 13Sep26 for Suite 115" in p4_text, "Monthly cleaning standard phrasing missing!"
assert 'Analyzing a 6-month sample history for GoGoMeds Select (E10747), this specific analyte "Sermorelin 0.9mg/ml" has had no prior Out-of-Specification (OOS) occurrences under USP <71> sterility testing.' in p4_text, "6-month history standard sentence missing!"
assert "ETX-260624-0206" in p4_text, "Sibling sample ETX-260624-0206 missing!"
assert "sample-to-sample cross-contamination is highly unlikely." in p4_text, "Cross-contamination conclusion missing!"
print("[PASS 6] Page 4 narrative fully verified against all 8 user instructions.")

# Check 6: Page 5 Narrative Audit
p5 = doc[4]
p5_text = " ".join(p5.get_text().split())
assert "Phase II investigation is initiated" not in p5_text, "Phase II sentence found in Page 5 narrative!"
assert "sample retesting is required" not in p5_text, "Retesting sentence found in Page 5 narrative!"
assert "Therefore, the likelihood of laboratory error contributing to this failure is considered minimal, and the original result is deemed valid." in p5_text, "Page 5 conclusion missing!"
print("[PASS 7] Page 5 narrative fully verified (no speculative Phase II / retest).")

# Check 7: Page 6 Checkboxes & Signatures
p6 = doc[5]
w_p6 = {w.field_name: w for w in p6.widgets()}
assert w_p6['Check Box87'].field_value == 'Off', "Check Box87 should be Off (no assignable lab cause)!"
assert w_p6['Check Box88'].field_value == 'Yes', "Check Box88 should be Yes (No, OOS is valid)!"
assert w_p6['Check Box89'].field_value == 'Off', "Check Box89 should be Off (Phase II not initiated)!"
assert w_p6['Text Field52'].field_value == 'Not Applicable', "Phase II justification should be Not Applicable!"
assert w_p6['Text Field53'].field_value == 'Qiyue Chen', "Investigator name missing!"
assert w_p6['Text Field54'].field_value == 'Robin Seymour', "Supervisor name missing!"
print("[PASS 8] Page 6 checkboxes and sign-off fields fully verified.")

# Check 8: Page 7 Table 1 & Table 2
p7 = doc[6]
p7_text = " ".join(p7.get_text().split())
assert "Table 1: Information for ETX-260624-0251 (BUD-T90) under investigation" in p7_text, "Table 1 title missing!"
assert "Bacillus megaterium" in p7_text, "Table 1 Bacillus megaterium missing!"
assert "2 x 200mL TSB (Original and Subculture Bottles)" in p7_text, "Table 1 media bottle details missing!"
assert "Table 2: Environmental Monitoring" in p7_text, "Table 2 title missing!"
assert "Personnel EM" in p7_text, "Personnel EM banner missing!"
assert "Biological Safety Cabinet EM for E001314" in p7_text, "BSC EM banner missing!"
assert "Weekly Active Air Sampling of Cleanroom Suite 115 (E001737) with Processing BSC" in p7_text, "Weekly Active Air banner missing!"
assert "Weekly Surface Sampling of Cleanroom Suite 115 (E001737) with Processing BSC" in p7_text, "Weekly Surface banner missing!"
assert "6 CFUs (ISO 8 115)" in p7_text, "6 CFUs missing in Table 2!"
assert "Micrococcus luteus (4 CFUs)" in p7_text, "Micrococcus luteus missing in Table 2!"
assert "Kocuria indica and Kocuria rhizophila (1 CFU)" in p7_text, "Kocuria missing in Table 2!"
assert "Pending FunITS (1 CFU, Hyphae)" in p7_text, "Pending FunITS missing in Table 2!"
print("[PASS 9] Page 7 Table 1 & Table 2 fully verified in Master PDF.")

doc.close()

# Check 9: Word Document Tables
wdoc = docx.Document(docx_path)
assert len(wdoc.tables) == 2, f"Expected 2 tables in Word doc, got {len(wdoc.tables)}"
tbl0 = wdoc.tables[0]
tbl1 = wdoc.tables[1]
print(f"Word doc Table 0 rows: {len(tbl0.rows)}, Table 1 rows: {len(tbl1.rows)}")
assert len(tbl0.rows) == 2, "Table 0 should have 2 rows (header + sample row)"
assert len(tbl1.rows) == 16, f"Table 1 should have 16 rows, got {len(tbl1.rows)}"

# Check uniform font size 7.0pt across all paragraphs in table cells
non_7pt_runs = 0
for t in [tbl0, tbl1]:
    for row in t.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    if r.font.size and abs(r.font.size.pt - 7.0) > 0.05:
                        print(f"Warning: Non-7pt run in cell: '{r.text}' has {r.font.size.pt} pt")
                        non_7pt_runs += 1
assert non_7pt_runs == 0, f"Found {non_7pt_runs} non-7pt font runs in Word tables!"
print("[PASS 10] Word document strictly complies with uniform 7.0pt font standard.")

print("\n" + "="*60)
print("100% COMPLETE AUDIT PASSED! READY FOR SUBMISSION!")
print("="*60)
