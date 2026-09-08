# Product Requirements Document: OOS Tools

## Overview
This project contains automated reporting tools for Eagle Analytical's Out-of-Specification (OOS) investigations, specifically focusing on Sterility Testing. The core functionality generates the "Phase I Summary" narratives and extracts data to populate PDF/Word OOS reports.

## Core Features
1.  **Smart Justification Engine (Automated Defense Mechanism):**
    *   **Goal:** Automatically build a defensive narrative when environmental monitoring (EM) yields positive microbial growth that is deemed unrelated to the test sample's contamination.
    *   **The 4-Step Shielding Mechanism:**
        1.  *Rule 1 (Not Identical):* Compare the microbial IDs from EM with the positive test sample. If they are different, state they are isolated events.
        2.  *Rule 2 (Physical Isolation):* Highlight that the sample manipulation occurred in an ISO 5 Primary Engineering Control, while weekly EM hits typically occur in the ISO 8 background environment.
        3.  *Rule 3 (Transfer Pathway Cutoff):* If analysts' daily glove/surface plates are clean, argue there was no viable transfer pathway from outer rooms to the ISO 5 BSC.
        4.  *Rule 4 (Macro-Environment Security):* If no other concurrently processed samples were positive, state that the testing environment was operating optimally and cross-contamination did not occur.
3.  **Environmental Monitoring (EM) Module (`em_logic.py`, `pages/EM.py`):**
    *   **Goal:** Standalone investigation module for OOS results originating on EM plates (Surface, Settling, Personnel/Glove, Cleanroom Air).
    *   **SOP Reference:** 2.600.002
    *   **Features:** Handles exceeded action/alert levels, organism identification, analyst interviews, and defensive transient contamination logic.

4.  **Standardized SKU / Module Generator Workflow (`create_new_module.py`):**
    *   Allows instant instantiation of new OOS test modules (templates, logic engine, and Streamlit UI page) in seconds using a single command.

## Finalized & Locked Features
*   **"RS Reviewed" Celsis Narrative Format:** The Celsis reporting logic (`pages/Celsis.py` and `celsis_logic.py`) has been completely overhauled to match the QA-approved "RS Reviewed" standard. This includes merging analyst introductions, specific phrasing for airflow (Suite 115B -> 115A -> 115), and separate detailed EM paragraphs for Processing and Aliquoting. **DO NOT modify this narrative flow without explicit user permission.**
*   **Smart Justification Verbiage:** The precise sentences used in the `smart_just` block (e.g., "Notably, the colony morphology...", "Also, the absence of contamination on analyst glove plates...") are finalized and locked based on user approval.
*   **Cross-Contamination & Lot History Logic:** Dynamic generation of text regarding sample-to-sample contamination and lot history checking is finalized.
*   **EM Module 5-Template Ecosystem & 7-Page PDF Pipeline:** EM OOS module templates (`EM OOS P1 template.docx`, `EM OOS P1 template 0.docx`, `EM OOS P1 template.pdf`, `tables for em.docx`, dynamic Page 7 Table PDF generator), `em_logic.py`, `pages/EM.py`, and `create_new_module.py` are live and fully verified against 7 real production G-drive EM reports.
*   **Scan RDI ZenQMS SOP Migration:** Scan RDI SOP references have been completely upgraded across all templates (`ScanRDI OOS template.docx`, `ScanRDI OOS template 0.docx`, `ScanRDI OOS P1 template.docx`, `ScanRDI OOS P1 template 0.docx`, `ScanRDI OOS template.pdf`, `ScanRDI OOS P1 template.pdf`) and Python generators (`pages/ScanRDI.py`) from legacy numbering (`2.600.023 / 2.700.004`) to the new ZenQMS standard: **`MICRO-SOP-12 (Rev 16, Effective 24Jul26)`** for Rapid Scan RDI Test using FIFU Method, and **`ENG-SOP-4 (Rev 05, Effective 24Jul26)`** for Scan RDI Operations and Maintenance. Section A tables, compliance rows 12/13, narrative paragraphs, and PDF form fields are locked to this standard.
*   **Scan RDI Multi-Analyst Interview Deduplication:** In `pages/ScanRDI.py` and batch processors, Section B `Analyst interviewed?` comment (`smart_comment_interview` / `Text Field13`) and Paragraph 1 narrative dynamically aggregate and deduplicate all involved analysts (Prepper, Processor, Changeover Processor, and Reader). Grammatically formats 1 analyst (`Yes, analyst X was interviewed comprehensively.`), 2 analysts (`Yes, analysts X and Y were interviewed comprehensively.`), or 3+ analysts with Oxford comma (`Yes, analysts X, Y, and Z were interviewed comprehensively.`). Smart variables are guaranteed to be passed to `final_data_docx` before Word rendering.
*   **Scan RDI FDA cGMP Assignable Cause Defense Standard:** In accordance with FDA Guidance for Industry (Aseptic Processing & Investigating Out-of-Specification Test Results for Pharmaceutical Production), initial sterility test positives may only be invalidated when contamination is unequivocally ascribed to laboratory error. When EM recoveries occur, the defense narrative strictly follows the "No assignable laboratory cause identified → no evidence linking laboratory conditions to the positive → original positive remains valid" framework:
    1. *Work Surface Integrity:* Zero recovery across all four ISO 5 BSC work surfaces.
    2. *Personnel Isolation:* Touch plate recoveries (e.g. *M. luteus* cocci) demonstrated to be non-matching in morphology to test sample isolates.
    3. *Desiccated / Inconclusive Settling Plate Defense:* When species ID cannot be definitively resolved due to plate desiccation, defend by establishing that *an organism-level microbiological match could not be established*, bracketed by negative ISO 5 work-surfaces and absence of aseptic breaches.
    4. *Weekly Facility EM Physical Segregation:* Air/surface hits in lower-classified ISO 7/8 background areas are defended by physical segregation and transport via disinfected, lidded bins.
    5. *Absence of Broader Contamination:* Comprehensive review confirming all concurrent samples processed on the same date tested negative with no clustering pattern.
    6. *FDA cGMP Conclusion:* Conclude that available EM data do not identify an assignable laboratory source and do not support laboratory-introduced contamination.
    7. *Two-Tier Investigation Architecture (Result Attribution vs Analytical Sterility):* When sample results are inadvertently transposed during post-analytical manual data entry/review, the investigation strictly separates the event into two distinct tiers: (a) Post-analytical documentation/transcription error reconciliation against original instrument data confirming true sample identities (`ETX-260804-0101` = actual pass, `ETX-260805-0189` = true fail); (b) Microbiological contamination evaluation for the true failing sample, maintaining strict demarcation between post-analytical documentation oversight and the absence of any laboratory analytical/testing cause.

## Pending/Future Work
*   **Roll out Smart Justification to USP <71>:** The engine is live for Celsis and Scan RDI, but `USP71.py` still needs its underlying logic updated to utilize the 4-Step Shielding Mechanism and the new "RS Reviewed" narrative format (adjusting for its specific workflow).
*   **Template Updates:** The underlying `.docx` and `.pdf` templates need manual layout updates (by the user in Word/Acrobat) to accommodate the significantly longer narrative text before they can be perfectly auto-filled without cutoff/font-shrinking.
