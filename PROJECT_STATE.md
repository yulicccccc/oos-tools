# Project State: OOS Tools

## Version: v12.4.21 (Previous: v12.4.20)

## Current Phase: Total Elimination of Acrobat '+' Overflow Icons on Form (9) Transcriptions (v12.4.21)

### What Was Completed
- **Root Cause Resolution for Acrobat '+' Overflow Icons**:
  - Identified the root cause of the '+' icons in OOS-262394: When populating the blank ZenQMS template CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (9).pdf, 56 text widgets had 	ext_fontsize = 0.0. PyMuPDF's appearance stream generator defaulted these to 12pt, causing text in constrained cells (height 15–20pt) to exceed bounding boxes and trigger Acrobat's '+' overflow indicator.
- **Universal Golden Font Inheritance**:
  - Updated 
ebuild_oos_262394_with_form9.py and generate_master_oos_262394.py to extract and apply the complete dictionary of exact font sizes directly from 	emplates/golden/ScanRDI_Golden_OOS-262289.pdf across all 61 form widgets:
    - Section A: Text Field1 (7.5pt), Text Field3 (6.0pt), Text Field4 (8.5pt), Text Field7 (7.5pt), Text Field8-10 (5.18pt).
    - Section B Comments: Text Field13 (4.825pt), Text Field14 (8.5pt), Text Field15/16 (9.0pt), Text Field17 (6.725pt), Text Field18 (4.825pt), Text Field21 (4.825pt).
    - Page 2 Reagents & Equip: Text Field22/23 (6.0pt), Text Field32/33 (4.4pt), Text Field43/44 (7.05pt/8.5pt).
    - Narratives: Text Field49 (7.0pt), Text Field50 (6.5pt), Text Field51 (11.0pt).
- **Zero-Overflow Automated Verification**:
  - Implemented regex stream inspection verifying min_y >= 0 across all page /AP/N streams.
  - Verified: **0 overflowing fields detected** (overflow_count == 0).
  - Image cropping visually confirmed complete elimination of '+' icons in Section A and Section B.
- **Desktop Deliverables**:
  - Saved clean, zero-overflow PDF to Desktop: CORP-FORM-21- 20 Oct 2026 OOS-262394 Optio RX (E18631) - ScanRDI (Fixed).pdf.

### Previous Releases
- **Integration of Concurrent OOS-262410 Defense into OOS-262414 & Single-Row Table Verification (v12.4.20)**



## Current Phase: Integration of Concurrent OOS-262410 Defense into OOS-262414 & Single-Row Table Verification (v12.4.20)

### What Was Completed
- **Concurrent OOS Investigation Cross-Referencing for OOS-262414**:
  - In OOS-262414 Section II Narrative (Text Field50), integrated the second sterility failure that occurred on the same testing date (22-Oct-2026):
    - Concurrent failure: Sample ETX-261019-0832 under **OOS-262410** (BelieveRX Pharmacy (E74120), Tirzepatide 10mg, Cyanocobalamin 0.5mg/mL, Lot 101626-7B).
    - Established clear cross-contamination defense:
      1. Physical separation: Processed in ISO 7 Cleanroom 144 (BSC E001937) vs. Cleanroom 145 (BSC E001938).
      2. Personnel separation: Processed by Halaina Smith (HS 5th) and read by Maryam Naeem (MN) vs. Karla Silva (KSM 12th) and read by Sonal Uprety (SU).
      3. Morphological distinction: Cocci and diplococci shaped morphology (OOS-262410) vs. rod-shaped morphology (OOS-262414).
      4. Zero shared laboratory vectors or cross-contamination pathways.
- **Strict Table Formatting & Single-Row Invariant (Rule 87 / RULE 0AE)**:
  - Maintained single-row layout without dividing sub-row borders for Weekly Active Air sampling in Table 2 across both Master PDF Page 7 and Standalone Tables Word/PDF.
  - All 4 active air recoveries (ETX-261027-0484, ETX-261027-0487, ETX-261027-0492, ETX-261027-0494) formatted vertically with EagleTrax hyperlinks.
- **SOP Version Synchronization (MICRO-SOP-12 Rev 17)**:
  - Retained verified authoritative SOP versions: MICRO-SOP-12 (17) Effective 28Sep26, Rev: 17 and ENG-SOP-4 (05) Effective 24Jul26, Rev: 05.
- **Internal OOS Tracking Log Row 222 Verification**:
  - Verified row 222 of sheet 'Scan RDI OOS' in scratch/Sterile Lab - OOS Tracking Log.xlsx accurately records OOS-262410 (BelieveRX Pharmacy, ETX-261019-0832, Lot 101626-7B, HS 5th, MN, L-Suite 144, cocci and diplococci).
- **Rule 0M Dual Deliverable Enforcement**:
  - Master 7-Page PDF recompiled and synchronized to Desktop: CORP-FORM-21- 22 Oct 2026 OOS-262414 Reshape Rx Houston (E74105) - ScanRDI.pdf.
  - Standalone Word table: Tables OOS-262414 Reshape Rx Houston (E74105) - ScanRDI.docx.
  - Exactly 2 canonical deliverables maintained per case on Desktop.

### Previous Releases
- **Verification of 11/1 Count, SOP Rev 17 Synchronization & OOS-262394 ZenQMS Alignment (v12.4.19)**

### What Was Completed
- **Confirmation of Ground-Truth Event Counts**:
  - Re-verified raw run data (`102026-2606-1.csv` and Scan Data Packet Page 73): Test sample `ETX-261017-0071` (`hs12`, Analysis ID `102026-2606-1034`) is confirmed as strictly **11 total events / 1 confirmed microbial event** exhibiting rod-shaped morphology.
- **Authoritative SOP Version Synchronization (MICRO-SOP-12 Rev 17)**:
  - Synchronized both `OOS-262394` and `OOS-262414` with the authoritative version from ZenQMS: `MICRO-SOP-12 (17)`, Effective Date `28Sep26`, `Rev: 17`, alongside `ENG-SOP-4 (05)`, Effective Date `24Jul26`, `Rev: 05`.
- **Concurrent OOS Investigation Cross-Referencing**:
  - In `OOS-262394` Section II Narrative (`Text Field50`), explicitly integrated the two concurrent positive cases tested on 20-Oct-2026:
    - Sibling positive `ETX-261019-0668` (`hs53`, cocci) under **`OOS-262395`** (CasaPharma Rx).
    - Third positive `ETX-261017-0024` (`pg36`, rod, BSC 1317) under **`OOS-262396`** (Link Pharmacy).
- **Authentic ZenQMS Form Re-transcription**:
  - Rebuilt Master 7-Page PDF `CORP-FORM-21- 20 Oct 2026 OOS-262394 Optio RX (E18631) - ScanRDI.pdf` onto the authentic blank template printed by Robin Seymour on 20-Oct-2026 at 11:53:02 PM UTC (`CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (9).pdf`).
  - Re-synchronized both canonical deliverables (`CORP-FORM-21-...pdf` and `Tables ...docx`) for both `OOS-262394` and `OOS-262414` on Desktop per Rule 0M.

### Previous Releases
- **Full Delivery of Scan RDI OOS-262414 & Single-Row Table Enforcement (v12.4.18)**

### What Was Completed
- **Full Deliverable Package for Scan RDI OOS-262414 (Reshape Rx Houston (E74105))**:
  - Sample `ETX-261021-0286` (Tesamorelin 20mg/2ml, Lot QC264093, Tested 22-Oct-2026).
  - Executed 1:1 clone of Golden Benchmark `templates/golden/ScanRDI_Golden_OOS-262289.pdf` Pages 1–6 with exact AcroForm slot substitutions.
  - Page 1: Sample ID `ETX-261021-0286`, Lot `QC264093`, Analyte `Tesamorelin 20mg/2ml`, Dosage Form `Injectable`, Tested Date `22-Oct-2026`, Read Date `22-Oct-2026`, Manager Notified Date `23-Oct-2026`, Target Technical Review Completion Date `09-Nov-2026`.
  - Pages 3–5: Comprehensive narrative integrating "No Method Suitability Test on File", processing order `KSM 12th` (`ksm12`), zero concurrent batch positives, negative controls, and positive control `A. brasiliensis` (Lot 05222026-16404-AB). Fully compliant with Rule 75.1 (no vague estimates) and Rule 75.2 (objective terminology).
  - Page 7: Standalone Table 1 (Information for ETX-261021-0286: 145 events, 84 confirmed events, bold italic *Rod-shaped morphology*) and Table 2 (Environmental Monitoring from Testing Performed on 22Oct26).
  - **Rule 87 / RULE 0AE (Single-Row Invariant)**: In Table 2, Weekly Active Air Sampling strictly housed within **ONE SINGLE ROW (无分层 / no horizontal sub-row dividing borders)**, separating all 4 recoveries, EagleTrax hyperlinks, and pending statuses vertically via cell paragraph breaks.
  - Generated Master 7-Page PDF: `CORP-FORM-21- 22 Oct 2026 OOS-262414 Reshape Rx Houston (E74105) - ScanRDI.pdf` (strictly 7 pages).
  - Generated Standalone Word Table: `Tables OOS-262414 Reshape Rx Houston (E74105) - ScanRDI.docx` (strictly 1 page).
  - Synchronized exactly the 2 canonical deliverables to Desktop per Rule 0M.
- **Internal OOS Tracking Log Update**:
  - Updated sheet `'Scan RDI OOS'` row 221 in `scratch/Sterile Lab - OOS Tracking Log.xlsx` and `scratch/Sterile_Lab_OOS_Tracking_Log_downloaded.xlsx`.

### Previous Releases
- **Narrative Alignment with Rule 75.1 & 75.2 & Deep Forensics Review (v12.4.17)**

### What Was Completed
- **Full Narrative Alignment with Rule 75.1 & Rule 75.2**:
  - Eliminated the unverified estimate "approximately 60 samples" in Page 4 cross-contamination evaluation, strictly replaced by verified sequence tracking: sample ETX-261017-0071 (hs12) and sibling positive ETX-261019-0668 (hs53) separated by 41 negative samples.
  - Replaced subjective modifier "distinctly different cellular morphologies" with compliant regulatory phrasing "different cellular morphologies".
  - Recompiled Master PDF (CORP-FORM-21- 20 Oct 2026 OOS-262394 Optio RX (E18631) - ScanRDI.pdf, strictly 7 pages) and synchronized to Desktop.
- **Deep Forensics Check & Open Questions Formulation**:
  - Identified 4 specific decisions/data points requiring user input (28-Sep SOP revision/effective date, ZenQMS blank naming gate, Scan Data Packet printed event count verification, and concurrent positive sample OOS ID).

### Previous Releases
- **Strict 2-File Deliverable Protocol Enforcement (Rule 0M)**:
  - Eliminated prohibited redundant deliverables (Tables OOS-262394 Optio RX (E18631) - ScanRDI.pdf and OOS-262394 Optio RX (E18631) - ScanRDI.docx) from Desktop and Git tracking.
  - Locked exactly the 2 canonical deliverables on OneDrive Desktop and Root:
    1. File 1: Master 7-Page PDF CORP-FORM-21- 20 Oct 2026 OOS-262394 Optio RX (E18631) - ScanRDI.pdf
    2. File 2: Standalone Tables Word Doc Tables OOS-262394 Optio RX (E18631) - ScanRDI.docx
  - Updated scratch/generate_master_oos_262394.py and scratch/generate_tables_oos_262394.py to prevent redundant copies from ever syncing to Desktop or Root.

### Previous Releases
- **Delivery of Scan RDI OOS-262394 Master Package & Tracking Log Extension (v12.4.15)**:

### What Was Completed
- **Full Deliverable Package for Scan RDI OOS-262394 (Optio RX (E18631))**:
  - Sample ETX-261017-0071 (Niacinamide/Tirzepatide 2mg/17.65mg/mL (4mL) Inj., Lot 3001002, Tested 20Oct26).
  - Executed 1:1 clone of Golden Benchmark 	emplates/golden/ScanRDI_Golden_OOS-262289.pdf Pages 1–6 with exact AcroForm slot substitutions, locking 67 widgets on Page 1, Text Field57 for header OOS ID, and perfect typography.
  - Generated official Master PDF CORP-FORM-21- 20 Oct 2026 OOS-262394 Optio RX (E18631) - ScanRDI.pdf (strictly 7 pages: Pages 1–6 AcroForm Form + Page 7 Standalone Table 1 & Table 2).
  - Generated official Master Word Document OOS-262394 Optio RX (E18631) - ScanRDI.docx.
  - Generated Standalone Word Tables Tables OOS-262394 Optio RX (E18631) - ScanRDI.docx and PDF Tables OOS-262394 Optio RX (E18631) - ScanRDI.pdf (strictly 1 page).
  - Synchronized all 4 deliverable files to Desktop (C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\) and workspace root (Documents\OOS\).
- **Enforcement of Active Air Multi-Recovery Single-Row Invariant (RULE 0AE / Rule 87 / 无分层)**:
  - Formatted all 4 weekly active air recoveries within **ONE SINGLE TABLE ROW** without horizontal sub-row dividing borders:
    1. 2 CFUs (ISO 7 144 Sec I) -> ETX-261027-0484 (Pending)
    2. 21 CFUs (ISO 8 143 Sec I) -> ETX-261027-0487 (Pending)
    3. 17 CFUs (ISO 8 143 Sec II) -> ETX-261027-0492 (Pending)
    4. 16 CFUs (ISO 8 142) -> ETX-261027-0494 (Pending)
  - All recoveries vertically aligned via blank paragraph breaks within columns 5, 6, and 7, preserving uniform 7.0pt Times New Roman typography and active hyperlinks.
- **Sterile Lab OOS Tracking Log Extension**:
  - Updated Row 218 in sheet 'Scan RDI OOS' of Sterile Lab - OOS Tracking Log.xlsx / Sterile_Lab_OOS_Tracking_Log_downloaded.xlsx:
    - ECD Date: 9-Nov-26
    - Assigned: QYC (ECD has been extended)
    - Status: Pending
    - Comments: 'Pending Micro ID ETX-261027-0487\nETX-261027-0492\nETX-261027-0494\nETX-261027-0484 as of 04Nov2026'
    - Full sample metadata: ETX-261017-0071, OOS-262394, 20-Oct-26, Optio RX (E18631), Niacinamide/Tirzepatide 2mg/17.65mg/mL (4mL) Inj., 3001002, HS, MN, Parallel Suite, 144, 
od shaped morphology, HS 12th.

### Previous Releases
- **Strict 1-to-1 Template Substitution from OOS-262080 (v12.4.11)**:
  - Directly cloned scratch/Celsis table OOS-262080.docx (backed up to .history/) to generate Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx, preserving 100% of its native Table Grid and w:tcBorders, exact column widths, cell margins, shading, underlined 9pt titles, and single-paragraph <w:br/><w:br/> multi-species formatting.
  - Eliminated 100% of self-invented extra sentences, repetitions, and visual cutoffs across Pages 1-6 by matching OOS-262080 paragraph-for-paragraph, field-for-field, and font-size-for-font-size.

### Previous Releases
- **Table Gridline Borders Restoration (v12.4.10)**:
  - Identified the root cause of borderless tables flagged by the user: python-docx `add_table()` defaults to normal table style without borders unless `<w:tblBorders>` is explicitly declared on `tblPr`.
  - Implemented `set_table_borders(table)` in `scratch/rebuild_oos262333_perfect.py` adding standard thin black borders (`single`, `sz="4"`, `000000`) across `top`, `bottom`, `left`, `right`, `insideH`, and `insideV`.
  - Applied `set_table_borders()` to Table 1 (`t1`), Table 2 (`tbl`), and Table 3 (`tbl`), completely eliminating the borderless floating text artifact.
  - Verified drawings stroke count in Master PDF increased from 117 to 706 on Page 7 and 92 to 619 on Page 8, matching benchmark `OOS-262080`.
  - Expanded automated test suite (`scratch/verify_celsis_tests.py`) to 50 assertions, strictly verifying `tblBorders` on all Word tables and drawing counts on PDF pages. All 50/50 PASSED (100%).
  - Recompiled and synchronized Master PDF (8 pages) and Standalone DOCX (2 pages) to Desktop and Documents.

### Previous Releases
- **100% Forensic Alignment with Golden Benchmark OOS-262080 (v12.4.9)**:
  - Conducted line-by-line forensic comparison against `OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - RS Reviewed - Signed by - QYC.pdf`.
  - **Page 1 Form Field Rectification**:
    - `Text Field2`: Restored to Sample ID (`ETX-261002-0288`), resolving previous field shift.
    - `Text Field4`: Restored to Product Description (`TESAMORELIN 8 MG/ML (3 ML) Injection`).
    - `Text Field5`: Restored to Dosage Form (`Injection`).
    - `Text Field6`: Restored to Lot Number (`LG342011076`).
    - `Text Field57`: Header OOS ID (`262333`) preserved and visible on Page 1 and across all pages.
    - `Check Box34` & `Check Box36`: Refrigerated storage marked strictly `Yes`, `N/A` strictly `Off` (never check both).
  - **Page 2 Reagents & Equipment**:
    - Cleanroom Suite: Corrected to official equipment ID `E001736 (CR 114)`.
    - Reagent Lots: ATP Positive Control lot strictly reconciled to `022601-1483` across both `Text Field22` and `Text Field25`.
    - Primary Engineering Controls: Included `ISO 5 BSC E001316 and ISO 5 BSC E001798` (Cal Due Jun 2027) alongside incubators `E001356` and `E001357`.
    - EM Conformity: `Check Box67` strictly `Yes`, `Check Box69` strictly `Off` (zero ambiguity).
  - **Page 3 GDP Sign-off & Phase I Options**:
    - `Text Field48`: Corrected to GDP sign-off format `'N/A QYC 22Oct26'`.
    - `Check Box84`: Strictly `Off`.
  - **Pages 3-5 Narrative Restoration & Fit**:
    - Restored full comprehensive multi-paragraph narratives (P3: 3386 chars, P4: 3828 chars, P5: 3863 chars).
    - Verified zero overflow with substantial positive clearances (P3 clearance 133.8pt, P4 clearance 245.2pt, P5 clearance 230.3pt).
  - **Tables Architecture**:
    - Table 1: Column 2 labeled strictly `Aliquoting Analyst` (`Cuong Du`).
    - Table 1 Microbial ID: Pure species *Bacillus paralicheniformis* without Gram stain notation.
    - Table 2 & Table 3: Section banner cells properly merged across all 9 columns.
    - Table 3 Bracketing: Strictly labeled `Date Before Aliquoting`, `Date of Aliquoting`, `Date After Aliquoting`.
    - Table 3 Row 13 Active Air: Pure 7 Oxford Nanopore sequenced species in italics separated by blank lines, without CFU counts and without `(Pending Microbial ID)` tags.
  - **Deliverable Formats & Verification**:
    - Master PDF: Strictly 8 pages (P1-6 Form, P7 Table 1 & Table 2, P8 Table 3).
    - Standalone Word: Strictly 2 pages (P1 Table 1 & Table 2, P2 Table 3).
    - Automated Test Suite (`scratch/verify_celsis_tests.py`): 45 / 45 test assertions PASSED (100%).
    - Synchronized across Desktop and Documents workspaces.

### Previous Releases
- **Full Migration to User's Canonical Master Table Grid Architecture (v12.4.8)**:
  - Replaced legacy/Mills table structure with the user's verified golden standard grid (derived from `build_usp71_tables_262329.py`):
    - **Table 1 (6 columns)**: `Processing Analyst` (`Alex Saravia`) | `Reading Analyst` (`Cuong Du`) | `Sample ID` (`ETX-261002-0288` with hyperlink) | `Related Microbial ID` (`ETX-261012-0654`) | `Media with microbial growth` (`1 x 300mL FTM`) | `Microbial ID` (*Bacillus paralicheniformis* / Gram-positive rod).
    - **Table 2 & Table 3 (9 columns)**: `Environmental Monitoring (EM) Sampling Site` (1.50 in) | `Frequency` (0.55 in) | `Date (DDMMMYY)` (0.60 in) | `Analyst` (0.50 in) | `Day /Week(s)` (0.85 in) | `Observation` (0.85 in) | `EM Plate ETX Number` (0.85 in) | `Microbial ID` (1.35 in) | `Notes` (0.45 in).
    - **4 Gray Section Dividers (`#D9D9D9`)**: `Personnel EM`, `Biological Safety Cabinet EM for E00XXXX`, `Weekly Active Air Sampling of Cleanroom Suite 114 (E001737) with Processing BSC`, `Weekly Surface Sampling of Cleanroom Suite 114 (E001737) with Processing BSC`.
    - **Table 3 Active Air Row 14**: Single-row invariant, `ETX-261021-0578` blue hyperlink, 7 Oxford Nanopore sequenced species in italic font separated by blank lines, without CFU counts.
    - **Seamless 2-Page Flow**: Table 1 and Table 2 flow naturally together on Page 1; Table 3 occupies Page 2 with explicit page break, completely eliminating margin clipping and floating title artifacts.
  - Recompiled Master PDF (`CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf`) and Standalone DOCX (`Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx`).
  - Synchronized across Desktop and Documents workspaces.

### What Was Completed
- **Elimination of Table 1 Border Clipping & Precise Margin Restoration (v12.4.7)**:
  - Identified the exact root cause of Table 1's left border clipping (Reading Analyst cut off): section margin was shrunk to 0.5" while Table 1 had -635 dxa negative indent, pushing table edge to X = 4.25pt.
  - Restored standard 1.0" (72pt) margins from QA benchmark 262316_MILLS EM TABLE.docx, placing Reading Analyst at X = 63.4pt with wide, pristine borders.

- **Elimination of Word Floating Table (w:tblpPr) Caption Inversion (v12.4.7)**:
  - Stripped all w:tblpPr floating attributes from Table 2 and Table 3, converting them to strict inline tables.
  - Replaced fragile manual paragraph spacers with explicit w:pageBreakBefore on Table 2 and Table 3 titles.
  - Set Table 2 and Table 3 left indent to -551 dxa (-27.55pt), perfectly positioning them from X = 50.0pt to X = 594.2pt inside the 612.0pt page.
  - Verified Table 2 caption at top of Page 8 (Y = 72.5pt) and Table 3 caption at top of Page 9 (Y = 72.5pt).

- **AcroForm Page 1 Text Field3 Typography & Complete Elimination of [+] Badge (v12.4.7)**:
  - Calibrated Text Field3 to 5.5pt in standard Helvetica, ensuring 0 line wrapping on (Celsis Sterility Processing Analyst) and 0 overflow.

- **Recompiled & Synchronized Deliverables (v12.4.7)**:
  - Standalone DOCX: Strictly 3 Pages (P1 Table 1, P2 Table 2, P3 Table 3).
  - Master PDF: Strictly 9 Pages (P1-6 Form, P7 Table 1, P8 Table 2, P9 Table 3).
  - Fully synchronized across Desktop and Documents workspaces.

- **Text Field3 Analyst Block Auto-Fit & Total Elimination of Overflow Plus Sign (v12.4.6)**:
  - Identified the exact root cause of the black [+] overflow indicator in Text Field3 (Name of Analyst who Performed the Test): setting 6.225pt was slightly too wide for the 100.48pt box width, causing (Celsis Sterility Processing Analyst) to wrap into two lines, which pushed the bottom reading analyst Cuong Du (CCD) out of the box.
  - Calibrated font size to 5.8pt in standard Helvetica (Helv), matching the non-wrapping single-line behavior of benchmark OOS-262316.
  - Verified 100% elimination of the [+] sign, with all 4 analysts and their role titles displaying completely on 8 clean lines.
  - Recompiled Master PDF and synchronized to desktop and documents workspace.

- **Standard Laboratory Typography & Single Canonical Deliverable (v12.4.5)**:
  - Fully restored form font to standard corporate Helvetica (Helv) matching Qiyue Chen and Robin Seymour's signed baseline OOS-262080, completely eliminating experimental Times-Roman overrides.
  - Re-calibrated standard font sizes across all fields (10pt for dates/names, 9.5pt for SOP/method headers, 8.5-9.0pt for narratives, 6.8pt for equipment, 4.8pt for reagents), ensuring zero + overflow symbols across P1-6.
  - Eliminated redundant hyphenated filename variant, standardizing strictly on the single canonical desktop file:
    CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf.

- **Standard TSB Culture Media Nomenclature Rule (v12.4.4 / Rule 89 / RULE 0AG)**:
  - Standardized all references to TSB culture media strictly to **Tryptic Soy Broth (TSB)** (or TSB).
  - Banned compendial aliases Soybean-Casein Digest Medium (TSB) and SCDM across all investigation forms, narratives, and tables to align 100% with Eagle Analytical SOPs (MICRO-SOP-44, MICRO-SOP-2) and QA-approved practice.
  - Re-verified Page 3 narrative in CORP-FORM-21 for OOS-262333, regenerated Master PDF, verified strictly 9 pages and zero overflow '+' signs, and synchronized both Desktop and Documents deliverables.

- **Deep Forensic Table Grid Calibration & Zero Text Duplication (v12.4.3)**:
  - Discovered that previous programmatic table generation caused duplicate cell text (e.g. `None\nNone`, `Date Before Testing\nDate Before Testing`) due to merging cells after population.
  - Cloned the exact XML table grid from QA-approved benchmark `262403_PGRrx EM TABLE.docx`, cleaned body spacing with explicit page breaks, mapped exact unique cell column indices, and regenerated both DOCX and PDF deliverables.
  - Verified 100% elimination of duplicate text, flawless layout across all 3 table pages, active hyperlinks on P7 and P9, and zero overflow `+` signs across P1-6.

- **Email Deep Penetration & Forensic Ground-Truth Comparison against QA-Approved Deliverables**:
  1. Connected to Outlook MAPI and systematically retrieved all QA-approved Celsis investigation packages signed by Quality Assurance (Tahsin Kevenk) and Microbiology Supervisor Robin Seymour (RS), including `OOS-262316` (Mills Pharmacy, Rev version post-QA review), `OOS-262403` (PGRrx LLC), and `OOS-262080` (Optimal Balance Pharmacy).
  2. Uncovered and eliminated all 10 systemic defects in the Celsis OOS package:
     - **Defect 1 (Page Count & Table Architecture)**: Master PDF was previously 8 pages with squeezed/missing tables. Corrected to **strictly 9 pages** (P1–6 AcroForm, P7 Table 1, P8 Table 2, P9 Table 3). Standalone Word table deliverable corrected to **strictly 3 pages** (P1 Table 1, P2 Table 2, P3 Table 3).
     - **Defect 2 (Table 1 Columns & Headers)**: Previous Table 1 mistakenly had Processing and Aliquoting Analysts as columns. Corrected strictly to QA standard: `Reading Analyst` (`CCD`), `Day of Incubation` (`Day 7`), `Sample ID` (`ETX-261002-0288` with hyperlink), `Related Microbial ID` (`ETX-261012-0654`), `Media with growth` (`1 x 300mL FTM`), and `Microbial ID` (*Bacillus paralicheniformis* / Gram-positive rod).
     - **Defect 3 (Page 1 Tested Date vs Read Date)**: `Date Field0` was previously set to `12-Oct-2026`, which contradicted the 7-day incubation period. Corrected strictly to Inoculation/Testing date **`05-Oct-2026`**! `Date Field1` (Read Date), `Date Field2` (OOS Initiated Date), and `Date Field3` (Manager Notified Date) set to **`12-Oct-2026`**.
     - **Defect 4 (Page 1 SOP Reference Fields)**: Corrected `Text Field8` strictly to **`MICRO-SOP-44`**, `Text Field9` strictly to **`03-Jan-2025`**, and `Text Field10` strictly to **`01`**.
     - **Defect 5 (Page 2 Primary Incubators Missing)**: Added primary Celsis incubators `Incubator E001356 (Monitored by Sensor E001450)` and `Incubator E001357 (Monitored by Sensor E001449)` (Cal Due: `Jan 2027
Feb 2027`) alongside backup incubators (`E001034`, `E001031`) and primary ISO 5 BSCs (`E001316`, `E001798`).
     - **Defect 6 (Page 2 Reagents Formatting)**: Formatted `Text Field22` and `Text Field23` with explicit `\r` newlines and clean sections (`Materials and Reagents`, `Celsis Reagents and Kits`, `Monthly Cleaning`).
     - **Defect 7 (Page 3 Check Box 84)**: Turned `Check Box84` **Off** (previously mistakenly checked Yes for chromatography re-dilution).
     - **Defect 8 (Table 3 Aliquoting Bracketing Labels)**: Set Table 3 bracketing labels strictly to **`Date Before Aliquoting`**, **`Date of Aliquoting`**, and **`Date After Aliquoting`** (addressing QA's explicit critique in OOS-262316).
     - **Defect 9 (Active Air Microbial ID Precedence & Purity)**: For Table 3 Row 14 (ETX-261021-0578), incorporated Oxford Nanopore sequencing results (*Moraxella osloensis*, *Sporisorium graminicola*, *Staphylococcus haemolyticus*, *Staphylococcus hominis*, *Kocuria indica*, *Ustilago maydis*, *Kocuria rhizophila*), separated by empty lines, without CFU counts and without `(Pending Microbial ID)` tags, italicized.
     - **Defect 10 (Page 5 QA Compliance Phrasing)**: Refined EM evaluation to explicitly state "absence of microbial recovery (No Growth) was demonstrated across all critical ISO 5 zones and personnel monitoring on the date of testing", and detailed physical airlock/barrier door segregation and positive pressure cascade isolating the outermost ISO 8 gowning room.
  3. **Verified Deliverables Synchronized**:
     - `CORP-FORM-21- 12-Oct-2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf` (Strictly 9 pages, 0 overflow symbols, active clickable hyperlinks on P7 and P9).
     - `Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx` (Strictly 3 pages, Table 1 on P1, Table 2 on P2, Table 3 on P3).
     - Synced to Desktop (`C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\`) and workspace root (`Documents\OOS\`).
  4. **Codified Rule 88 / RULE 0AF**:
     - Fully documented in `PRD.md` and `.agents/AGENTS.md` (both in top pinned summary and full rule section).

## Previous Phase: OOS-262394 Standalone Investigation Tables Completed (Single-Row Active Air Layout)

### What Was Completed
- **Scan RDI OOS-262394 Standalone Tables Generation & Single-Row Layout Invariant (Rule 87 / RULE 0AE)**:
  1. Enforced single-row layout without sub-layering for multiple active air recoveries.
  2. Verified data for Optio RX (E18631), Sample ETX-261017-0071, 11 total events, 1 confirmed microbial event.
  3. Generated strictly 1-page deliverables and codified Rule 87.
