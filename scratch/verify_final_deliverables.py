import os, fitz, docx

onedrive_desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

pdf_33 = os.path.join(onedrive_desktop, "OOS-262133 EM SMO 143 Air 04SEP2026.pdf")
pdf_34 = os.path.join(onedrive_desktop, "OOS-262134 EM SMO 142 Air 04SEP2026.pdf")
docx_33 = os.path.join(onedrive_desktop, "EM table OOS-262133 04SEP2026.docx")
docx_34 = os.path.join(onedrive_desktop, "EM table OOS-262134 04SEP2026.docx")

print("Checking existence of required files...")
for p in [pdf_33, pdf_34, docx_33, docx_34]:
    assert os.path.exists(p), f"Missing: {p}"
    print(f"  EXISTS: {os.path.basename(p)} ({os.path.getsize(p)} bytes)")

print("\nChecking absence of prohibited files (Rule 0M)...")
for bad_f in ["EM table OOS-262133 04SEP2026.pdf", "EM table OOS-262134 04SEP2026.pdf"]:
    bp = os.path.join(onedrive_desktop, bad_f)
    assert not os.path.exists(bp), f"PROHIBITED FILE FOUND ON DESKTOP: {bp}"
    print(f"  CLEAN: No {bad_f}")

print("\nChecking PDF pages and text integrity...")
for pdf_p, oos_id in [(pdf_33, "262133"), (pdf_34, "262134")]:
    doc = fitz.open(pdf_p)
    print(f"\n--- Checking {os.path.basename(pdf_p)} ---")
    print(f"  Total pages: {len(doc)} (Must be 7)")
    assert len(doc) == 7, f"Expected 7 pages, got {len(doc)}"
    
    # Page 3
    p3_txt = doc[2].get_text()
    assert "please see attached" in p3_txt, "Page 3 bracketing text missing!"
    print("  Page 3 text: COMPLETE (contains 'please see attached')")
    
    # Page 4
    p4_txt = doc[3].get_text()
    assert "failed sterility testing" in p4_txt, "Page 4 sterility testing text missing!"
    assert "30 Aug 2026" in p4_txt, "Page 4 monthly cleaning missing!"
    print("  Page 4 text: COMPLETE (contains 'failed sterility testing' and '30 Aug 2026')")
    
    # Page 5
    p5_txt = doc[4].get_text()
    assert "compliance with existing procedures" in p5_txt, "Page 5 conclusion missing!"
    print("  Page 5 text: COMPLETE (contains 'compliance with existing procedures')")
    
    # Page 6
    p6_txt = doc[5].get_text()
    assert "Qiyue Chen" in p6_txt, "Page 6 Initiator missing!"
    assert "Robin Seymour" in p6_txt, "Page 6 Supervisor missing!"
    print("  Page 6 text: COMPLETE (contains 'Qiyue Chen' and 'Robin Seymour')")
    
    # Page 7 (Table)
    p7_txt = doc[6].get_text()
    assert "Read Dates and Incubation Observation" in p7_txt, "Table 1 missing!"
    assert "Environmental Monitoring for Analyst and Cleanroom" in p7_txt, "Table 2 missing!"
    assert "Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs" in p7_txt, "Active air banner missing!"
    assert "Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs" in p7_txt, "Surface banner missing!"
    assert "&" not in p7_txt, "Prohibited '&' found in Table text!"
    print("  Page 7 text: COMPLETE (Robin QA standards, zero '&')")

print("\nALL VERIFICATIONS PASSED 100%!")
