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

## Pending/Future Work
*   **Roll out Smart Justification to Scan RDI and USP <71>:** The engine is live for Celsis, but `ScanRDI.py` and `USP71.py` still need their underlying logic updated to utilize the 4-Step Shielding Mechanism and the new "RS Reviewed" narrative format (adjusting for their specific workflows).
*   **Template Updates:** The underlying `.docx` and `.pdf` templates need manual layout updates (by the user in Word/Acrobat) to accommodate the significantly longer narrative text before they can be perfectly auto-filled without cutoff/font-shrinking.
