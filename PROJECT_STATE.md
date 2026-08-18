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
