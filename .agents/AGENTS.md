# Project-Scoped Rules for Eagle Analytical OOS Tools

## SOP: Creating a New OOS Module (Template Handoff Workflow)

Whenever the user initiates the creation of a new OOS module (e.g., adding a new test like USP71), follow this strict Human-AI collaborative SOP for the **Template Handoff (Step 1)**.

### The Contract
The core architecture of this system relies on `docxtpl` (Jinja2 syntax in Word). The Word document is the ultimate source of truth for all static text, formatting, and tables. Python code MUST NOT hardcode any long boilerplate text; it only populates variables marked with `{{variable_name}}`.

### Step 1 Workflow (Human & AI Responsibilities)

1. **Human Action**: 
   - The user duplicates an existing `.docx` template (e.g., `Celsis OOS P1 template 0.docx`) and renames it to the new test (e.g., `USP71 OOS P1 template 0.docx`).
   - The user manually updates the `{{}}` variable tags inside the Word document to match the specific needs of the new test (e.g., changing `{{celsis_result}}` to `{{usp71_id}}`).
   - The user **MUST CLOSE** Microsoft Word to release the file lock.
   - The user notifies the AI to proceed.

2. **AI Action**:
   - **Global Replace**: The AI runs a Python script (using `python-docx`) to globally search and replace the old test name (e.g., "Celsis") with the new test name (e.g., "USP <71>") across all paragraphs and tables in the new `.docx` file.
   - **Syntax 体检 (Sanity Check)**: The AI runs a Python script (using `re` and `python-docx`) to scan the document for malformed Jinja2 tags (e.g., spaces inside tags like `{{  subculture _name  }}`). The AI should automatically fix these syntax errors to prevent `docxtpl` from crashing.
   - **Variable Extraction (The Contract)**: Once the document is clean, the AI uses `docxtpl.DocxTemplate(filepath).get_undeclared_template_variables()` to extract the definitive list of all `{{}}` variables.
   - **Handoff**: The AI presents this alphabetized variable list to the user in the chat. This list acts as the absolute data contract for Step 2 (writing the `_logic.py` engine).

### Behavior Enforcement
- DO NOT attempt to modify a `.docx` file while the user has it open (it will throw `Permission denied`). Instruct the user to close the file first.
- NEVER skip the syntax check phase. MS Word formatting and typos often introduce invisible characters or spaces into `{{}}` tags which will fatally crash `docxtpl`.

---

## 🚨 ARCHITECTURAL CONSTRAINT: The Dual-Template Bulk Insertion

This architecture requires TWO parallel templates due to differing rendering constraints between Word and PDF:

### 1. `template.docx` (Standard Word Template)
- **Purpose**: The primary Word document for generation.
- **Rule**: It uses a single massive macro tag `{{ smart_phase1_summary }}` to ingest the entire narrative string from the backend logic in one go.

### 2. `template 0.docx` (PDF-Compatible Hollow Template)
- **Purpose**: The fallback created strictly as a workaround for `pypdf` limitations (PDF AcroForms have character limits on single fields).
- **Rule**: It splits the massive narrative into two giant placeholder macros: `{{ smart_phase1_part1 }}` and `{{ smart_phase1_part2 }}`.

### Python Backend Obligation
- **DO NOT** attempt to map granular English sentences inside the Word templates.
- **MUST** assemble complete, multi-paragraph text blocks (including domain-specific logic) entirely within the Python `_logic.py` engine.
- **MUST** push this massive variable block simultaneously to `smart_phase1_summary` (for Word) and `smart_phase1_part1/2` (for PDF).

---

## 🚨 Template Preservation & File History Rule

To prevent data loss and preserve user-created assets, you MUST follow these instructions:
1. **Never Overwrite Templates Directly**: Before making any modification or running scripts on template files (like `template 0.docx`, `template.docx`, `template.pdf`), you MUST create a copy of the original file in a `.history/` directory or rename it with a timestamp suffix (e.g. `_backup_YYYYMMDD_HHMMSS`) to preserve the historical version.
2. **Preserve User Assets**: Treat all user-made template documents as sacred. Never run bootstrap or tag-fixing scripts that overwrite them unless the user explicitly commands you to do so.

---

## 🚨 EM OOS Investigation Standards & Golden Rules (QA / Neha Feedback)

Whenever drafting or reviewing an Environmental Monitoring (EM) OOS investigation report:

1. **Table 1 is Mandatory**: Always include Table 1 (Read Dates & Incubation Observation) alongside Table 2. Never submit Table 2 alone or omit Table 1.
2. **Weekly EM Bracketing Scope (Week of Testing Only)**: For EM OOS reports, only include the **week of testing** (active air & surface sampling). Do NOT pull in the prior week's weekly EMs unless explicitly requested.
3. **Initiator Field**: On Page 1 (`Text Field0`), the Initiator must ALWAYS be the actual analyst who initiated the OOS event in ZenQMS (e.g., `Simin Mohammad`), NEVER the drafting analyst or reviewer (`Qiyue Chen`).
4. **Incubator E001031 & E001034 Calibration Due Dates**: The calibration due dates for Incubators E001031 and E001034 in 2026 are **Aug 2027** (do NOT write August 2026).
5. **Surface Sampling Negative Control Statement**: For surface sampling EM OOS reports, always add this statement at the end of the second paragraph on Page 5:
   `"Additionally, it is important to note that no growth was observed on the other three surface sampling plates from the date of testing."`
6. **Consumables (Contact Plate / TSA Plate) Lot & Expiration**: Accurately verify the Contact Plate / TSA Plate lot number and expiration date against the lab dispensing logs.
7. **Action Level Wording ("Met" vs. "Exceeded")**:
   - If the Action Level is $\ge 1\text{ CFU/Plate}$ (or $\ge X$) and the count is exactly $1\text{ CFU}$ (or $X$), write **"met the action level"**, NEVER "exceeded the action level".
   - Only use "exceeded" if the count is strictly greater than the numerical threshold.
8. **No Speculative Laboratory Error Statement (cGMP Root Cause Standard)**:
   - Do NOT state that an event "could likely be attributed to an inadvertent laboratory error" unless a specific, documented, and verifiable laboratory discrepancy was identified during testing or interview.
   - If no error was found, state that the recovery was an isolated, transient event, the analyst adhered to approved aseptic protocols with no deviations, the critical environment remained in control, and no specific or assignable laboratory discrepancy was identified.
9. **Table 1 Setup Date vs. Sampling Date Consistency**:
   - The setup date in Table 1 must strictly match the actual test/sampling date of the plate (e.g., `11MAY 2026`), not an earlier weekly monitoring date (e.g., `07MAY 2026`).


## Side-by-Side Diff Comparison Rule (左右对照找茬式对比规范)
**CRITICAL**: Whenever you present text revisions, draft updates, QA feedback responses, prompt refinements, or document modifications to the user:
1. **Always Use Side-by-Side (左右对照) Comparison Tables**:
   - **NEVER** output vertically stacked "Before" followed by "After" large text blocks.
   - You **MUST ALWAYS** format text modifications into a side-by-side Markdown comparison table (类似“大家来找茬”).
2. **Strict Table Column Layout**:
   - Column 1: `模块 / 关键要点 (Section / Focus)`
   - Column 2: `🔴 修改前 (Original / Before)`
   - Column 3: `🟢 修改后 (Revised / After)`
   - Column 4: `💡 差异解析与核心改动 (Key Changes & Rationale)`
3. **Line-by-Line Alignment & Granular Correspondence (逐段逐句严格对齐)**:
   - Ensure the rows correspond directly line-by-line or sentence-by-sentence, so the user can easily trace every word, phrase, and logical alteration across the left and right columns.
   - Use bolding or highlight markers (`**...**`) in both the Before and After columns to instantly flag specific additions, deletions, and phrasing shifts.
4. **Spot-the-Difference Visual Ergonomics (大家来找茬极致体验)**:
   - Make the contrast crystal clear so the user never has to scroll up and down or guess where subtle changes occurred.


## 🚨 Laboratory Physical Topology, EM Dual-Track Tracing & cGMP Investigation Golden Rules

### 1. Physical Spatial Topology Chain (物理空间拓扑链与房间推导法则)
- **Do not blindly accept scan documents without spatial validation.** Every BSC is physically situated in a specific cleanroom, which belongs to a specific cleanroom suite:
  $$\text{BSC 1316} \longrightarrow \text{Cleanroom 114B (ISO 7 核心室)} \longrightarrow \text{Cleanroom 114A (ISO 7 缓冲室)} \longrightarrow \text{Anteroom 114 (ISO 8 走廊)} \longrightarrow \text{Suite 114}$$
  $$\text{BSC 1798} \longrightarrow \text{Cleanroom 114A (ISO 7 缓冲室)} \longrightarrow \text{Anteroom 114 (ISO 8 走廊)} \longrightarrow \text{Suite 114}$$
- **Weekly EM Scope Invariance**: Because testing in BSC 1316 is located in Cleanroom 114B (Suite 114), its Weekly Active Air & Surface monitoring **MUST STRICTLY belong to Suite 114**. Never accept or insert records from Suite 115 or other suites, even if inadvertently provided in raw scans.

### 2. EM Dual-Track Tracing Rule (EM 追溯“双轨制”法则：人跟人走，物跟物走)
- **Personnel EM (人身绑定 / 人跟人走)**:
  - Glove fingertip touch plates monitor the individual analyst's aseptic gowning and touch behavior.
  - Must follow the specific analyst's personal schedule and shift history across days.
  - *Example*: Aliquoting was performed by CCD on 08Sep26 in BSC 1798. His bracketing personnel plates must be traced to CCD's own shift history: Pre = 03Sep26 (CCD), Test = 08Sep26 (CCD), Post = 15Sep26 (CCD).
- **Settling & Surface EM (设备与空间绑定 / 物跟物走)**:
  - Settling plates and surface contact plates monitor the physical ISO 5 Critical Zone of the Biological Safety Cabinet.
  - Must follow the continuous operational history of that specific BSC, regardless of who operated it.
  - *Example*: For BSC 1798 bracketing: Pre = 07Sep26 (ALA), Test = 08Sep26 (CCD), Post = 09Sep26 (ALA).

### 3. Raw Bench Records Precedence (穿透系统看板凳 / 原始纸质台账至上原则)
- In cGMP compliance, the true analyst is the person who performs the bench-level aseptic manipulation, disinfects the hood, plates the samples, and signs the physical paper logbooks.
- System users who merely change status, upload files, or enter electronic test results in EagleTrax / ZenQMS (e.g. `aalanis`) are Data Coordinators / Submissions Coordinators, NOT the Aliquoting Analyst.
- Table 1 Aliquoting Analyst must be the physical bench operator (Cuong Du / CCD).

### 4. Zero Hallucination & Sub-threshold Recovery Defense (ALCOA+ 零臆想与微量检出闭环防守)
- **Zero Hallucination (不能臆想，实事求是)**: If a record is missing or not located in the scan, state clearly that it was not found. Never invent dates, analysts, or negative results.
- **Sub-threshold Recovery Accountability**: If a weekly EM plate yields recovery (e.g., 1 CFU in ISO 8 Anteroom 114 on 04Sep26 under `ETX-260914-0487` by ISS, and 1 CFU on 10Sep26 under `ETX-260921-0520` by SMO), it MUST be accurately reported in the bracketing tables.
- **Airtight Defense Triad (合规防守闭环三部曲)**:
  1. *Physical Segregation & Pressure Differential*: The recovery occurred in the outermost ISO 8 Anteroom (114), which is physically segregated from the ISO 7 buffer/cleanrooms (114A/114B) and ISO 5 BSCs, maintained by positive pressure cascading inward-to-outward.
  2. *Containment & Transport*: Samples and media are transferred in disinfected, lidded bins on carts without open atmospheric exposure to anteroom air.
  3. *Zero Critical Zone Breach*: 100% absence of microbial recovery (0 CFU / No Growth) across all analyst glove touch plates, settling plates, and ISO 5 BSC work surfaces throughout testing and aliquoting stages.

### 5. Defensive Post-Processing in Document Automation (自动化工程后置防御性校验)
- Word templates may contain hardcoded static text in table cells (e.g., static `SMO` in weekly rows instead of Jinja2 tags).
- Automation scripts must perform surgical run-level inspection and overwrites using `python-docx` to ensure the final delivered document matches ground truth data without breaking table formatting.
