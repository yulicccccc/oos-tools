# Project State: OOS Tools

## Current Phase: Rollover to New Project Environment

### What Was Completed
- The Celsis reporting logic (`celsis_logic.py`, `pages/Celsis.py`) has been fully redesigned to match the narrative structure of the "RS Reviewed" model (OOS-261165).
- A 4-Step Smart Justification Engine (Shielding Mechanism) has been developed and integrated into the Celsis module, successfully synthesizing defenses based on ID mismatch, physical isolation, transfer pathways, and macro-environment history.
- The underlying decision trees (flowcharts) for all three test types (Celsis, Scan RDI, USP <71>) have been finalized and documented in `OOS_Justification_Flowcharts.md`.
- Analyzed 7 real EM OOS PDF reports from G: drive and built the standalone **Environmental Monitoring (EM)** module (`em_logic.py`, `pages/EM.py`, `EM OOS P1 template.docx`, `EM OOS P1 template 0.docx`, `EM OOS P1 template.pdf`).
- Created a standardized **SKU Module Generator Workflow** (`create_new_module.py`) to automatically instantiate future OOS test modules in seconds.

### Current File Structure
The codebase has been successfully copied from the legacy directory to the new clean context boundary:
`C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS`

Key active files:
- `PRD.md`: Defines product requirements and locked features.
- `PROJECT_STATE.md`: This file, documenting current progress and next steps.
- `em_logic.py` & `pages/EM.py`: Environmental Monitoring (EM) module logic and Streamlit UI.
- `create_new_module.py`: Automated SKU workflow generator for adding new test modules.
- `celsis_logic.py` & `pages/Celsis.py`: Celsis integration.
- `scanrdi_logic.py` & `pages/ScanRDI.py`: Need updates to match the new engine.
- `usp71_logic.py` & `pages/USP71.py`: Need updates to match the new engine.

### Which Files Should NOT Be Touched
- Do NOT revert the "RS Reviewed" formatting in `celsis_logic.py` or `pages/Celsis.py`.
- Do NOT change the verbiage in the smart justification engine without explicit permission.

### Next Phase Goal
1. Wait for the user to manually adjust the blank PDF and Word templates to accommodate the new 6-page report length.
2. Roll out the Smart Justification Engine to `ScanRDI.py` and `USP71.py` in the new Project environment.

### Known Issues
- The current `.pdf` and `.docx` templates used by the Python scripts are still formatted for a 4-page report. Running the generator now will cause text to be cut off or shrink until the user provides the expanded 6-page templates.
