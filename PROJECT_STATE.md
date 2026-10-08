# Project State: OOS Tools

## Current Phase: Rollover to New Project Environment

### What Was Completed
- The Celsis reporting logic (`celsis_logic.py`, `pages/Celsis.py`) has been fully redesigned to match the narrative structure of the "RS Reviewed" model (OOS-261165), and has been subsequently updated based on the 2026-09-16 RS feedback for OOS-261878 (adding Reading Analyst tracking, "viable but not culturable" terminology for no-growth, and finalized conclusion).
- A 4-Step Smart Justification Engine (Shielding Mechanism) has been developed and integrated into the Celsis module, successfully synthesizing defenses based on ID mismatch, physical isolation, transfer pathways, and macro-environment history.
- The underlying decision trees (flowcharts) for all three test types (Celsis, Scan RDI, USP <71>) have been finalized and documented in `OOS_Justification_Flowcharts.md`.
- Analyzed 7 real EM OOS PDF reports from G: drive and fully built the standalone **Environmental Monitoring (EM)** 5-template ecosystem:
  1. `tables for em.docx`: Table 1 (Read Dates & Incubation Observation) + Table 2 (Bracketing Table).
  2. `EM OOS P1 template.docx`: Complete standard Word report template with Form 3.100.019.F01 + EM Table 1 & Table 2.
  3. `EM OOS P1 template 0.docx`: Dual-template split narrative version with EM Table 1 & Table 2.
  4. `EM OOS P1 template.pdf`: 6-page Form 3.100.019.F01 (157 AcroForm fields, 100% matched to production PDFs).
  5. Page 7 PDF Table Generator: Dynamically renders Table 1 & Table 2 with ReportLab and merges with the 6-page form to produce a complete 7-page official PDF report.
- Upgraded `em_logic.py` and `pages/EM.py` to support automatic 7-page PDF generation, Smart Paste parsing, and full Word/PDF rendering.
- Created a standardized **SKU Module Generator Workflow** (`create_new_module.py`) to automatically instantiate future OOS test modules in seconds.
- Migrated all Scan RDI SOP references across all Word (`ScanRDI OOS template.docx`, `template 0.docx`, `P1 template.docx`, `P1 template 0.docx`) and PDF templates to ZenQMS standards: `MICRO-SOP-12 (Rev 16, Effective 24Jul26)` and `ENG-SOP-4 (Rev 05, Effective 24Jul26)`.
- Fixed Scan RDI Section B `Analyst interviewed?` comment (`smart_comment_interview` / `Text Field13`) and paragraph 1 narrative to aggregate and deduplicate all involved analysts (`Prepper`, `Processor`, `Changeover`, and `Reader`), correctly producing comprehensive multi-analyst interview phrasing (e.g. `Yes, analysts Elysse Nioupin, Sonal Uprety, and Varsha Subramanian were interviewed comprehensively.`). Passed smart variables to `final_data_docx` prior to docx rendering.
- Implemented FDA cGMP-aligned "Assignable Cause" defense narrative engine in `pages/ScanRDI.py`, `scan_logic.py`, and `scratch/process_oos_261814.py`. The narrative establishes that initial sterility test positives cannot be invalidated without unequivocal proof of laboratory error, robustly defending the laboratory via 0 ISO 5 BSC work surface recovery, non-matching touch plate morphology, desiccated/unmatched settling plate isolates, physical cleanroom segregation, and negative concurrent sample testing without clustering. Generated and verified both Word and PDF reports for OOS-261814.
- Integrated Two-Tier Investigation Architecture into `scratch/process_oos_261814.py` for OOS-261814 (`GoGoMeds Select (E10747)`, sample `ETX-260805-0189`), documenting post-analytical result transposition (Analyst VV transcription error, Reviewer OA review oversight) vs verified instrument-generated data (`ETX-260804-0101` = actual pass, `ETX-260805-0189` = true fail) seamlessly inserted before Table 2 Environmental Monitoring, with full Word and PDF report generation.
- Resolved Scan RDI Table 2 duplication and missing Notes: upgraded `tables for scan.docx`, `ScanRDI OOS P1 template.docx`, `pages/ScanRDI.py`, and `scratch/process_oos_261814.py` with dynamic single-BSC row deduplication (removing redundant Surface and Settling changeover rows when `bsc_id == chgbsc_id`), removed date from BSC bracketing header as requested, widened Notes column by >2.6x to 1.11 inches, dynamically mapped Jinja2 Notes extraction (`{{ note_pers }}`, `{{ note_sett }}`, etc.), and dynamically removed the Trend Table (Table 3) whenever past OOS records <= 3.
- Implemented automated Hyperlink Preservation and Word-to-PDF conversion for Standalone Tables: Table 1 Sample ID now supports dynamic OpenXML `<w:hyperlink>` embedding (Eagle Trax submission URL `https://etrax.eagleanalytical.com/Submission/Details/...`). Standalone tables are exported to pixel-perfect 1-page PDF via Word COM, preserving active clickable hyperlink annotations.
- Integrated same-day downloaded QA Form (CORP-FORM-21 v11.1) transcription and 7-page PDF assembly pipeline for Scan RDI: Automatically transcribes complete OOS datasets into freshly downloaded ZenQMS forms (populating 65 standard boilerplate compliance fields alongside sample data). Enforced Table 2 Notes formatting standard (Capitalized with no trailing period: `Plate was desiccated on 5 day read`). Appended the standalone Table PDF as Page 7 to create the unified 7-page official OOS investigation PDF report.
- Successfully finalized Celsis OOS-262080 (ETX-260828-0527, Optimal Balance Pharmacy) complete package, resolving all 4 hidden traps in scan files (Suite 115 trap vs Suite 114 ground truth, ISO 8 114 weekly air 1 CFU recovery under ETX-260914-0487 on 04Sep26 by ISS, Aliquoting Analyst attribution to Cuong Du / CCD, and complete BSC 1316/1798 surface logs). Produced and synced 5 clean deliverables: Main DOCX, Standalone Tables DOCX, Main PDF, Standalone Tables PDF, and Complete Combined PDF package.
- Documented and locked the 5 core cGMP investigation rules into `.agents/AGENTS.md` and `PRD.md` (Physical Spatial Topology Chain, EM Dual-Track Tracing, Raw Bench Records Precedence, Zero Hallucination & Sub-threshold Recovery Defense, and Defensive Post-Processing).
- Transcribed finalized Celsis OOS-262080 report into official ZenQMS template `CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf`: Populated all 157 AcroForm fields (100% fidelity, 0 mismatches, 30 active check boxes). Applied golden calibrated font size (`9.2pt`) to Page 4 (`Text Field50`) and Page 5 (`Text Field51`) to account for the taller ZenQMS corporate header, achieving zero bottom cutoff and zero artificial bottom void. Backed up base form to `.history/` and assembled complete 8-page deliverable package with standalone tables.
- Injected active clickable EagleTrax hyperlink into Table 1 on Sample ID `ETX-260828-0527` (`https://etrax.eagleanalytical.com/SubmissionTest/Details/jzraMOOTFYUFD%24erkwgubw__`) across `Celsis table OOS-262080.docx`, `Celsis table OOS-262080.pdf`, `OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.docx`, and synchronized into all combined 8-page deliverable PDFs (`QYC.pdf`, `Complete.pdf`, and `(2).pdf`) preserving native Word-to-PDF OpenXML annotations.
- Fully integrated live Microbial Identification results from EagleTrax for both environmental monitoring plates into OOS-262080 deliverables:
  1. `ETX-260921-0520` (Aliquoting week active air plate, 10Sep26, SMO): Identified as `Micrococcus luteus` (`Gram (+) cocci`), with clickable link to `https://etrax.eagleanalytical.com/SubmissionTest/Details/qnwcLQO5BWeJhWCBG7jl8Q__` in Table 3.
  2. `ETX-260914-0487` (Processing week active air plate, 04Sep26, ISS): Identified as `Corynebacterium ureicelerivorans` (`Gram (+) short rods`) & `Mycobacterium grossiae` (`Gram (+) rods`), with clickable link to `https://etrax.eagleanalytical.com/SubmissionTest/Details/fd3G2StZClcy1TP2ES6BLw__` in Table 2.
  3. Master narrative (Page 5 `Text Field51` & Word docx narrative) updated with both precise species and Gram stain descriptions, verifying that the text fits naturally within the bounding box with ~28.4pt bottom margin, zero text overflow, and zero artificial padding.
  4. All deliverables compiled, verified, and synchronized across Desktop and repository paths.

- **Celsis OOS-262017 Table Finalization (Revive Rx Pharmacy, E00927):**
  1. Incorporated user's manual census updates resolving all previously pending TBD daily bracketing rows in Table 2 and Table 3 to `"No growth"` / `"N/A"` / `"None"`.
  2. Table 3 Row 13 (Weekly Active Air on 04Sep26) reconciled with QA RS review: analyst attributed to `ISS`, colony count reported as `2 CFU (ISO 8 114)` (reconciled for two isolates), embedded with native clickable hyperlink to `ETX-260914-0487`, and microbial identification formatted as *Corynebacterium ureicelerivorans* & *Mycobacterium grossiae* (italicized species, zero `&`).
  3. Generated pixel-perfect 2-page DOCX and PDF deliverables on Desktop and scratch with 100% cell centering and visual verification.
- **Celsis OOS-262017 Official ZenQMS Form Transcription (CORP-FORM-21 v11.1):**
  1. Backed up blank official template `CORP-FORM-21 - P1 31 Aug 2026.pdf` to `.history/CORP-FORM-21 - P1 31 Aug 2026_backup_20261007_115410.pdf`.
  2. Transcribed complete dataset from user's finalized OOS PDF into official ZenQMS form, incorporating user's manual update on Page 3 `Text Field48`: `"N/A QYC 07Oct26"`, initiating analyst Cuong Du, and clearing manager field for live signature.
  3. Harmonized 3-page narrative layout across Pages 3, 4, 5 (7 / 5 / 8 paragraphs) with calibrated font sizes (`8.35pt` / `8.5pt` / `8.5pt`), achieving 100% text visibility with zero Acrobat `+` clipping and zero text overflow.
  4. Appended 2-page standalone tables (`Celsis table OOS-262017.pdf`) with 4 native clickable hyperlinks to create the complete 8-page unified packet (`OOS-262017 ... (Complete).pdf` and `CORP-FORM-21 - P1 31 Aug 2026.pdf`).
  5. Synchronized all deliverables to Desktop and repository paths.

- **USP <71> OOS-262098 Master Production Package (Solyn LLC, E75000):**
  1. Completed standalone Table 1 and Table 2 (`Tables OOS-262098 Solyn LLC (E75000) - USP71.docx` / `.pdf`): strictly 1 page, Times New Roman 7.0pt, active clickable hyperlinks to EagleTrax submission URLs for `ETX-260902-0505`, `ETX-260910-0290`, and `ETX-260914-0487`.
  2. Incorporated verified lab scheduling rule: skipped weekend (`05Sep26` Sat & `06Sep26` Sun) and Labor Day holiday (`07Sep26` Mon) to trace the following processing date to Tuesday `08Sep26` (ES).
  3. Fully reconciled Weekly Active Air EM recovery on `04Sep26` from PRD: attributed to analyst `ISS`, colony count reported as `2 CFU (ISO 8 114)`, plate `ETX-260914-0487` (with active clickable link), and identified isolates formatted as *Corynebacterium ureicelerivorans* & *Mycobacterium grossiae* (both italicized, zero `&`).
  4. Fully integrated definitive molecular identification results for positive TSB bottle under `ETX-260910-0290`: *Microbacterium sp. PM5* (`Gram (+) rods`) from both TSA and SDA subcultures (reconciling preliminary presumptive Gram (-) broth reading).
  5. Generated full Master Word Report (`OOS-262098 Solyn LLC (E75000) - USP71.docx`) with embedded clean tables.
  6. Transcribed complete dataset into official ZenQMS form `CORP-FORM-21 - P1 04 SEP 2026.pdf` (157 fields, 30 checkboxes, Page 2 Incubator blank line rule alignment, calibrated font sizes via PyMuPDF). Appended Standalone Tables as Page 7 to create the unified 7-page official OOS investigation PDF report (`OOS-262098 Solyn LLC (E75000) - USP71.pdf` and `... - QYC.pdf`).
  7. Synchronized all deliverables across Desktop and Documents, and committed to GitHub.

- **USP <71> OOS-262098 Phase II Full Investigation Package (Solyn LLC, E75000):**
  1. Conclusively confirmed Product Contamination root cause: Initial test `ETX-260902-0505` failed on Day 6 in TSB (*Microbacterium sp. PM5*, `ETX-260910-0290`, Gram (+) rods). Retest `ETX-260914-0470` replicated exact kinetics on Day 6 in TSB with sequencing under `ETX-260921-0498` confirming identical strain ***Microbacterium sp. PM5*** (`Gram (+) short rods`). Refuted preliminary analyst handling error hypothesis and confirmed inherent lot contamination.
  2. Incorporated non-uniform microbial distribution scientific defense for prior passing ScanRDI test (`ETX-260807-0602`, no method suitability on file).
  3. Generated Standalone 1-Page Tables (`Tables OOS-262098 Solyn LLC (E75000) - Phase II.docx`/`.pdf`), Master Word Report (`OOS-262098 Solyn LLC (E75000) - Phase II.docx`), and Official 6-Page PDF Form (`OOS-262098 Solyn LLC (E75000) - Phase II.pdf` & `... - QYC.pdf`).
  4. Aligned Phase 1 (Form 3.100.019.F01) Page 6 and Phase II (Form 3.100.019.F02) dispositions and synced to Desktop and Documents.

### Current File Structure
The codebase is actively operating in the clean context boundary:
`C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS`

Key active files:
- `PRD.md`: Defines product requirements and locked features.
- `PROJECT_STATE.md`: This file, documenting current progress and next steps.
- `em_logic.py` & `pages/EM.py`: Environmental Monitoring (EM) module logic and Streamlit UI (7-page pipeline).
- `tables for em.docx`: Standardized EM Table 1 & Table 2 Word template.
- `EM OOS P1 template.docx`, `EM OOS P1 template 0.docx`, `EM OOS P1 template.pdf`: EM core templates.
- `create_new_module.py`: Automated SKU workflow generator for adding new test modules.
- `celsis_logic.py` & `pages/Celsis.py`: Celsis integration.
- `scanrdi_logic.py` & `pages/ScanRDI.py`: Need updates to match the new engine.
- `usp71_logic.py` & `pages/USP71.py`: Need updates to match the new engine.

### Which Files Should NOT Be Touched
- Do NOT revert the "RS Reviewed" formatting in `celsis_logic.py` or `pages/Celsis.py`.
- Do NOT change the verbiage in the smart justification engine without explicit permission.
- Do NOT alter the EM 5-template structure without explicit permission.

### Next Phase Goal
1. Roll out the Smart Justification Engine to `ScanRDI.py` and `USP71.py` in the new Project environment.
2. User manual layout polish on any specific template aesthetics if desired.
