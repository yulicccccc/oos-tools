with open('PROJECT_STATE.md', 'r', encoding='utf-8') as f:
    content = f.read()

target = '## Current Phase: USP <71> OOS-262329 Phase I & Phase II Completion'
idx = content.find(target)
assert idx != -1

new_header = '''## Version: v12.3.12 (Previous: v12.3.11)

## Current Phase: Supervisor Robin Seymour (RS) OOS-262329 Review Forensic Diff & Rule 86 Expansion Complete

### What Was Completed
- **Supervisor Robin Seymour (RS) OOS-262329 Comprehensive Forensic Comparison & Error Elimination**:
  1. **Exhaustive PDF & Table Diff Analysis**:
     - Completed side-by-side comparison between our drafted versions and RS reviewed versions (`OOS-262329 GoGoMeds Select (E10747) - USP71 - Phase I - Reviewed.pdf` & `Phase II - Reveiewed.pdf`).
     - Identified and codified all 10 key supervisory improvements:
       1) **Subculture Day 18+ Mandatory Table 3**: Dedicated Table 3 documenting subculture analyst AOD's 07Oct26 touch plates, BSC E001314 EM, and 06Oct26 weekly active air/surface data.
       2) **Cleanroom EM Recovery Rationale & Physical Segregation**: ISO 7 Room 115A recovery (1 CFU Corynebacterium segmentosum under ETX-261013-0297) transparently disclosed, and defended via physical airlock segregation, positive pressure cascading, and zero critical ISO 5/processing room 115B recovery.
       3) **Subculture Day Count Mathematical Consistency**: Subculture aliquot day count strictly calculated as 18 - 14 = 4 (`Day 4 of the subculture aliquots`), eliminating the previous `Day 5` error caught by the user.
       4) **Section E GDP Sign-off Standard**: Every fillable N/A in Section E strictly signed off as `N/A QYC 04Nov26` (never bare `N/A`).
       5) **Table 1 Scope Separation**: Phase I Table 1 restricted strictly to initial failing sample `ETX-260624-0251 (BUD-T90)`; Phase II Table 1 covers dual samples (initial + retest `ETX-261021-0445`).
       6) **Comprehensive Equipment Traceability**: Suite 115 all active incubators (`E001356/E001450`, `E001357/E001449`, `E001034/E001501`, `E001031/E001505`) and ISO 5 BSC `E001314` fully listed on Page 1/Page 2.
       7) **Day 14 Inconclusive Read Date Alignment**: Unified initial inconclusive read date and subculture date to the exact same calendar day (`07Oct26`).
       8) **Microbiology Checkbox Invariants**: Phase I Page 3 `Check Box79` checked `Yes` (N/A for re-injection); Phase II Page 2 `Check Box42` checked `Yes` (No for formula worksheet).
       9) **Narrative Compartmentalization**: Clean uppercase subheadings in Section D for initial vs subculture EM.
       10) **Phase I Conclusion & BUD Retest Transition**: Explicitly links valid initial failure to client request for BUD-T90 retest under Phase II.
  2. **Rule 86 Codified into PRD.md & .agents/AGENTS.md**:
     - Formally updated Rule 86 with all 10 golden rules to permanently prevent supervisory revisions in future sterility OOS cases.
  3. **Standalone Word Tables Document Updated**:
     - Updated `scratch/build_usp71_tables_262329.py` and regenerated `Tables OOS-262329 GoGoMeds Select (E10747) - USP71.docx` on Desktop to include Table 1, Table 2, and Table 3.
  4. **Signed Reviewed Deliverables on Desktop**:
     - `OOS-262329 GoGoMeds Select (E10747) - USP71 - Phase I - Reviewed - QYC.pdf`
     - `OOS-262329 - GoGoMeds Select (E10747) - Phase II - Reveiewed - QYC.pdf`
     - Both PDFs have been digitally signed by Qiyue Chen and are ready for supervisor approval routing.

'''

idx_ver = content.find('## Version:')
content_new = content[:idx_ver] + new_header + content[idx:]
with open('PROJECT_STATE.md', 'w', encoding='utf-8') as f:
    f.write(content_new)

print('Updated PROJECT_STATE.md successfully.')
