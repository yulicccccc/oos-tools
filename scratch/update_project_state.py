with open("PROJECT_STATE.md", "r", encoding="utf-8") as f:
    text = f.read()

new_section = """## Version: v12.4.23 (Previous: v12.4.22)

## Current Phase: Section B Material Discrepancy Checkbox Correction (Check Box37=Yes for Concurrent Run OOS) (v12.4.23)

### What Was Completed
- **Corrected Concurrent Run OOS Checkbox (`Check Box37`)**:
  - Addressed user feedback regarding Page 1 Section B: *"Are the other test results from this run under investigation?"*
  - Previously, `Check Box38` was checked ('No').
  - Corrected to **`Check Box37 = 'Yes'`** and **`Check Box38 = 'Off'`**:
    - Ground-truth rationale: On the date of testing (20-Oct-2026), sample `ETX-261019-0668` (`OOS-262395`, CasaPharma Rx) was processed in the **exact same testing batch and session** (`102026-2606-1`) by processing analyst Halaina Smith (`hs53`, following `hs12`).
    - Since other test results from this run were positive and under active investigation (`OOS-262395`), this item MUST strictly be marked **Yes**.
- **Deliverables Regeneration & Verification**:
  - Re-executed `rebuild_oos_262394_with_form9.py` and `generate_master_oos_262394.py`.
  - Zero-Overflow Verification: **0 overflowing fields detected** across all form pages.
  - Page count: Strictly 7 pages (6 pages AcroForm + 1 page Table 1/2).
  - Synchronized directly to Desktop: `CORP-FORM-21- 20 Oct 2026 OOS-262394 Optio RX (E18631) - ScanRDI.pdf`.

### Previous Releases
- **4-Analyst Roster Correction (Prepping: NT, Changeover: ALA) & Golden SOP Alignment in OOS-262394 (v12.4.22)**"""

split_point = text.find("### Previous Releases")
rest = text[split_point:]

header = "# Project State: OOS Tools\n\n"
new_text = header + new_section + "\n" + rest[len("### Previous Releases"):]

with open("PROJECT_STATE.md", "w", encoding="utf-8") as f:
    f.write(new_text)

print("Updated PROJECT_STATE.md to v12.4.23!")
