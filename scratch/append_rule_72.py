# Append Rule 72 to PRD.md and RULE 0Z to .agents/AGENTS.md

rule_72_prd = """

---

### Rule 72: Automated Pre-Flight Integrity & Anti-Contradiction Assertion Gate (`assert_no_facility_or_event_contradictions`)

**CRITICAL**: To permanently eliminate cross-document and cross-section contradictions (such as the event count mismatch or BSC room mismatch caught during OOS-262289 review), all generation pipelines and transcription scripts MUST call and pass `assert_no_facility_or_event_contradictions()` in `utils.py` before exporting:

1. **Event Count Single-Truth Invariant**:
   - The number of `total events` appearing anywhere in Page 3 narrative (`Text Field49`) MUST be 100% identical to Table 1 and the raw Scan Data packet (`expected_events`).
   - If any discrepancy is detected (e.g., text has `15 total events` while Table 1 has `19`), generation is immediately aborted with a `ValueError`.

2. **BSC & Cleanroom Spatial Invariant (Figure 6 Ground Truth)**:
   - `BSC E001938` is permanently locked to **`Cleanroom 145 (CR145)`**. Any clause pairing `BSC E001938` with `Cleanroom 144` triggers an immediate hard block.
   - `BSC E001937` is permanently locked to **`Cleanroom 144 (CR144)`**. Any clause pairing `BSC E001937` with `Cleanroom 145` triggers an immediate hard block.

3. **Active Air Morphology Defense Invariant (RULE 0X.4)**:
   - When environmental isolates include *Staphylococcus* or other cocci-shaped bacteria, claiming "morphology distinctly different from cocci" is strictly prohibited and automatically blocked.
"""

rule_0z_agents = """

## 🚨🚨 RULE 0Z (TOP PRIORITY): Automated Pre-Flight Integrity & Anti-Contradiction Assertion Gate (`assert_no_facility_or_event_contradictions`)

**CRITICAL**: Every OOS report generator and transcription script MUST enforce the automated assertion gate defined in `utils.py`:

1. **Event Count Single-Truth Invariant**:
   - `Text Field49` (Page 3 narrative) event count MUST match Table 1 (`expected_events`) exactly. Zero tolerance for narrative stating 15 while Table 1 states 19.

2. **BSC & Cleanroom Spatial Invariant (Figure 6 Ground Truth)**:
   - `BSC E001938` $\to$ **`Cleanroom 145 (CR145)`** ONLY. Pairing with `Cleanroom 144` is blocked by code.
   - `BSC E001937` $\to$ **`Cleanroom 144 (CR144)`** ONLY. Pairing with `Cleanroom 145` is blocked by code.

3. **Active Air Morphology Defense Invariant**:
   - Prohibits morphology comparison claims when environmental isolates are cocci.
"""

with open('PRD.md', 'ab') as f:
    f.write(rule_72_prd.encode('utf-8'))
print("Appended Rule 72 to PRD.md")

with open('.agents/AGENTS.md', 'ab') as f:
    f.write(rule_0z_agents.encode('utf-8'))
print("Appended RULE 0Z to .agents/AGENTS.md")
