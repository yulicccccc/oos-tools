# Project State: OOS Tools

## Current Phase: Rollover to New Project Environment

### What Was Completed
- The Celsis reporting logic (`celsis_logic.py`, `pages/Celsis.py`) has been fully redesigned to match the narrative structure of the "RS Reviewed" model (OOS-261165).
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
