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

### 6. Mandatory EM Logbook Census Protocol & Anti-Extrapolation Rule (台账严谨清点与零外推防陷阱法则)
- **The 6-Document Complete EM Census Checklist**: Whenever an EM scan PDF is received, the AI MUST systematically verify the presence of all 6 standard microbiology logbooks before declaring any EM investigation dataset complete:
  1. `3.600.002.F01` (Settling Sampling Logbook): Passive microbial settling in ISO 5 Critical Zone (BSC/LAFW).
  2. `3.600.002.F02` (Surface Sampling Logbook): Contact plates 1–4 on ISO 5 BSC interior surfaces.
  3. `3.600.002.F03` (Personnel Sampling Logbook): Operator left and right fingertip touch plates.
  4. `3.600.002.F04` (Anteroom/Buffer Room Surface Sampling Logbook): ISO 7 buffer and ISO 8 anteroom contact plates.
  5. `3.600.002.F05` (Active Air Sampling Logbook): Cleanroom volumetric active air monitoring plates.
  6. `3.600.018.F05` (BSC Cleaning & Disinfection Logbook): Daily disinfection and contact time verification.
- **Strict Anti-Extrapolation Mandate (F02 ≠ F01)**: Having `F02` (Surface) does NOT mean `F01` (Settling) exists or is negative. NEVER extrapolate across logbooks. If `3.600.002.F01` (Settling) is absent from the provided PDF scan, the AI MUST NOT assume "No Growth" or extrapolate from surface plates. It MUST explicitly flag the missing document and mark it as `[Pending]` in all draft tables and narratives until physical scans are provided.
- **Pre-Response Verification Gate**: When asked "还缺什么吗？" (Is anything missing?), systematically verify every page's header block against the 6-document checklist before answering.

### 7. Universal Table 1 Clickable Hyperlink Standard (Table 1 原生交互超链接常态化规范)
- Across ALL OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `EM`), whenever Table 1 is generated (in standalone table documents or master report Page 7/8 attachments), the Sample ID (`ETX-XXXXXX-XXXX`) **MUST ALWAYS** be embedded with an active, clickable external hyperlink directly pointing to its EagleTrax test details URL (`https://etrax.eagleanalytical.com/SubmissionTest/Details/...` or `/Submission/Details/...`).
- **OpenXML Native Hyperlink Injection**: In `python-docx`, construct an XML relationship `<w:hyperlink r:id="...">` with `w:rStyle="Hyperlink"` and blue underline formatting (`#0000FF`).
- **Word-to-PDF Interactive Preservation**: When converted via Word COM, the interactive link is compiled into the native PDF Annotation tree, ensuring QA reviewers and management can open the live EagleTrax sample page with a single click directly from any PDF reader.
- **Mandatory Execution**: Never render the Sample ID as plain unlinked text if the submission URL or ETX ID is known.

### 8. Automated EagleTrax Authentication & Differential Staining Protocol (EagleTrax 自动化查询与 EM 超链接规范)
- **Persistent SSO Session Re-use**: The AI leverages the persistent Playwright profile at `LOCALAPPDATA/pastdue_playwright_session` to bypass repetitive interactive logins.
- **Direct Live Extraction**: When an EM plate has microbial recovery, the AI autonomously navigates to its EagleTrax `#TestDetails` tab, scrapes the Differential Staining results (e.g. `Gram (+) rods`, `Gram (+) cocci`), captures screenshot evidence, and feeds the findings directly into the report.
- **EM Plate Hyperlink Invariance**: In Table 2, all positive EM Plate ETX numbers (e.g. `ETX-260908-0584` and `ETX-260908-0580`) **MUST ALWAYS** be embedded with clickable hyperlinks directly to their test details URL.
- **Defensive Temporary File Export**: Always compile through intermediate files (`temp_tables_export.pdf`) so desktop file locks (Acrobat/Edge) do not crash generation.

### 9. Page 2 Incubator & Equipment Blank Line Standard (Page 2 培养箱与设备空行对齐规范)
- In Page 2 `Text Field43` (Equipment ID) and `Text Field44` (Calibration Due Date), each individual incubator entry **MUST ALWAYS** be separated by an empty line (`\r \r`).
- The Sensor ID sits on the same line as the Incubator ID (e.g., `Incubator E001356 (Sensor E001450)`).
- The corresponding Calibration Due Dates in `Text Field44` **MUST** maintain identical 1-to-1 line spacing / empty lines (`\r \r`) so that every date block aligns horizontally with its respective incubator block.
- Standard pattern:
  - Left column (`Text Field43`):
    ```text
    Incubator E001356 (Sensor E001450)\r \r
    Incubator E001357 (Sensor E001449)\r \r
    Incubator E001034 (Sensor E001501)\r \r
    Incubator E001031 (Sensor E001505)
    ```
  - Right column (`Text Field44`):
    ```text
    Jan 2027 / Feb 2027\r \r
    Jan 2027 / Feb 2027\r \r
    Aug 2027 / Feb 2027\r \r
    Aug 2027 / Feb 2027
    ```

### 10. AcroForm PDF Checkbox Strict Whitelist & Exclusivity Rule (PDF 表单复选框严格白名单排他互斥规范)
- In PDF AcroForms, never rely on default/blank states in underlying templates.
- For every checkbox row (e.g. Yes / No / N/A), explicitly set the target selection to `'Yes'` and **ALL alternative options in that row to `'Off'`**.
- Never allow multiple check boxes in the same question row to be active simultaneously. Validate programmatically after compilation that each question row has exactly 1 box checked.

### 11. Natural Human Typography & Acrobat '+' Overflow Elimination Standard (原生手打字号与杜绝 '+' 截断留白规范)
**CRITICAL**: Across **ALL Eagle PDF documents and templates** (`ScanRDI`, `Celsis`, `USP <71>`, `Environmental Monitoring (EM)`, etc.), generated PDFs must look 100% like natural human manual input into Adobe Acrobat:
1. **Calibrated Auto-Fit Font Sizing vs. PyMuPDF 12pt Trap (针对紧凑表单字段设定适配字号，规避 PyMuPDF 12pt 陷阱)**:
   - **Root Cause of Acrobat `+` Icon**: When a multiline field is set to `text_fontsize = 0.0` (Auto), PyMuPDF's `/AP` appearance stream generator defaults to `/Helv 12 Tf` (12pt font). In constrained table cells (e.g., Section B Comments with height 15.8pt), 12pt text pushes lines 2 and 3 outside the bounding box, forcing Adobe Acrobat to display a `+` symbol and clip visible text to line 1.
   - **The Compliance Solution**: Multi-line fields in tight table cells MUST be assigned their calibrated auto-fit font sizes (matching Foxit/Acrobat auto-scale values from signed reference reports like `OOS-261877`):
     - Section B Comments: `Text Field13` (4.825pt), `Text Field14` (8.5pt), `Text Field15/16` (9.0pt), `Text Field17` (6.725pt), `Text Field18` (4.825pt), `Text Field19/20` (9.0pt), `Text Field21` (5.85pt).
     - Analyst Box (`Text Field3`): 6.225pt (all 4 analysts fit cleanly).
     - Incident Description (`Text Field7`): 8.0pt (3 lines fit cleanly).
     - Reagents & Facilities (Page 2): `Text Field22` (4.35pt), `Text Field23` (4.75pt), `Text Field32/33` (4.825pt), `Text Field43/44` (7.05pt).
   - In Adobe Acrobat, these sizes fit 100% of the text inside the cell, completely eliminating the `+` button and showing all content cleanly.
2. **Zero Artificial Bottom Padding & Trailing Line Breaks (杜绝人为底部留白与尾部空行)**:
   - **NEVER** append trailing carriage returns or blank lines to form field strings (e.g. `'Sample Name\r \r \r \r'`, `'Celsis ATP Positive Control\r \r'`, or trailing returns in text summaries).
   - Form field strings must end naturally and cleanly without artificial trailing spacing.
3. **Natural Text Flow in Narrative Sections (叙述段落自然延展)**:
   - For multi-paragraph narratives (Page 3 `Text Field49`, Page 4 `Text Field50`, and Page 5 `Text Field51`), maintain the golden calibrated font size `9.2pt` (or `10.3pt` in balanced layouts). The text fills the available box space naturally with ~20-30pt bottom margin, completely avoiding bottom voids or text truncation.
4. **Universal Invariance (全模块一律适用)**:
   - This rule is permanent and universally locked for all existing and future OOS automation workflows. No script or agent may deviate from it.

### 12. Cleanroom Facility Monthly Cleaning Schedule & Prior Bracketing Standard (洁净室月度深度清洁排班与最近追溯规范)
**CRITICAL**: In microbiology cleanroom facility operations and OOS investigation reporting (under `MICRO-SOP-9`: Cleaning and Disinfecting Procedure for Microbiology), cleanroom monthly cleaning and disinfection (using $H_2O_2$ vapor / chemical indicators across ISO 8 Anteroom, ISO 7 Buffer, ISO 7 Cleanroom, and ISO 5 BSCs) is conducted on an official bi-weekly Sunday schedule: **Mid-Month Sunday (月中周日)** and **Last Sunday (月末周日)**.

#### 1. Official 2026 Facility Monthly Cleaning Schedule (2026 官方排班表)
| Month (月份) | Mid-Month Sunday (月中周日) | Last Sunday (月末周日) |
| :--- | :--- | :--- |
| **August 2026** | August 16, 2026 (`16 Aug 2026`) | August 30, 2026 (`30 Aug 2026`) |
| **September 2026** | September 13, 2026 (`13 Sep 2026`) | September 27, 2026 (`27 Sep 2026`) |
| **October 2026** | October 11, 2026 (`11 Oct 2026`) | October 25, 2026 (`25 Oct 2026`) |
| **November 2026** | November 15, 2026 (`15 Nov 2026`) | November 29, 2026 (`29 Nov 2026`) |
| **December 2026** | December 13, 2026 (`13 Dec 2026`) | December 27, 2026 (`27 Dec 2026`) |

#### 2. Prior Bracketing Date Selection Rule (最近回溯与就近判定铁律)
- When drafting the narrative for any OOS investigation (e.g., Page 4 `Text Field50` / Page 5 `Text Field51` or Word report narrative):
  `"Monthly cleaning and disinfection of the outermost ISO 8 Anteroom... was performed on [DATE], as per MICRO-SOP-9..."`
- The `[DATE]` **MUST ALWAYS** be the single most recent monthly cleaning date that has already occurred strictly prior to or on the processing/testing date ($\text{cleaning\_date} \le \text{test\_or\_process\_date}$).
- **Both Mid-Month Sunday and Last Sunday must be evaluated**:
  - If processing/testing date is before the Mid-Month Sunday of month $M$, trace back to the **Last Sunday of month $M-1$**.
  - If processing/testing date is on or after Mid-Month Sunday but before Last Sunday of month $M$, trace back to the **Mid-Month Sunday of month $M$**.
  - If processing/testing date is on or after Last Sunday of month $M$, trace back to the **Last Sunday of month $M$**.
- *Concrete Examples*:
  - Processed on `01Sep26` $\to$ Most recent cleaning was `30 Aug 2026` (Last Sunday of Aug).
  - Processed on `08Sep26` $\to$ Most recent cleaning was `30 Aug 2026` (Last Sunday of Aug; Sep 13 had not yet occurred).
  - Processed on `14Sep26` $\to$ Most recent cleaning was `13 Sep 2026` (Mid-Month Sunday of Sep).
  - Processed on `28Sep26` $\to$ Most recent cleaning was `27 Sep 2026` (Last Sunday of Sep).
  - Processed on `15Oct26` $\to$ Most recent cleaning was `11 Oct 2026` (Mid-Month Sunday of Oct).
  - Processed on `26Oct26` $\to$ Most recent cleaning was `25 Oct 2026` (Last Sunday of Oct).

#### 3. Strict Pre-Processing Invariance & Anti-Post-Event Rule (绝对前置因果与严禁后置追溯铁律)
- **绝对禁止使用 Process 之后发生的清洁**：
  - 洁净室月度深度清洁是实验开展前的**准入前置保障条件 (Baseline Prerequisite Condition)**。
  - 样本在接种（Process）时所处的洁净室环境状态，只能由**接种时刻之前已经发生并确认合格**的深度清洁来保障。
  - 接种之后发生的清洁属于“未来事件 (Future Event)”，在时间线与因果律上与该样本当时的处理过程完全脱节，绝对禁止作为调查依据！
- **杜绝“月份字面直觉陷阱” (Avoid Month-Name Cognitive Trap)**：
  - **常见致命错误**：样本于 9 月初（如 `01Sep26` 或 `08Sep26`）接种处理，粗心的分析员或 AI 往往习惯性去翻看 9 月份的清洁记录（`13 Sep 2026` 或 `27 Sep 2026`），并误将 9 月中/末的日期写入报告。
  - **铁律判定**：在 9 月 1 日或 9 月 8 日当天，9 月 13 日和 27 日**根本尚未发生**！必须毫不犹豫地回溯到严格早于处理日的最近一次清洁——即 **`30 Aug 2026`**！
- **杜绝“前后括弧外推陷阱” (No Post-Cleaning Bracketing)**：
  - 虽然 ISO 5 关键区的人身手套和沉降碟采用“实验前、实验中、实验后”三点括弧式追踪（Pre, Test, Post）；
  - 但月度深度清洁**绝不存在“后置（Subsequent）月度清洁”这种论证用法**！报告中必须 100% 且唯一引用**处理前最近**的那一次已完成清洁。

#### 4. Universal Invariance (全模块一律适用)
- This schedule and prior bracketing rule apply universally across **ALL OOS modules** (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring (EM)`).

### 13. Mandatory 6-Month Client Sample History Census Protocol (客户 6 个月历史记录严谨审查与防臆想门禁规范 - 🚨 重大错误警示与永久铁律)
**CRITICAL**: In any laboratory OOS investigation (under FDA Phase I & Eagle OOS SOP requirements across `ScanRDI`, `Celsis`, `USP <71>`), an analysis of the client's past 6-month historical sample and testing performance is a mandatory regulatory component of the investigation report narrative (e.g., Page 5 `Text Field51` or Word report narrative):
`"An analysis of the six-month sample history for [Client Name] ([Client ID]) indicates that Eagle Analytical processed [N] samples for [Test Method] with [no / X] prior occurrences of an out-of-specification or positive result for this analyte during this period."`

#### 1. 🚨 血的教训与致命错误警示 (Lesson Learned & Permanent Ban)
- **ZERO HALLUCINATION & FATAL MISTAKE (曾经发生的重大失误，绝不再犯！)**:
  - 曾经在起草 OOS 报告时，漏问客户 6 个月送检历史，或直接使用没有确凿数据支撑的模糊空话（如 "processed samples with no prior occurrences"）。
  - 在 cGMP / FDA 审计标准下，任何未经验证的历史陈述均属严重合规漏洞！如果客户实际存在历史不合格批次，调查报告中随意定论将构成严重的合规造假隐患！
  - **现已永久固化为全模块铁律：严禁在未穿透核实真实数据前私自草拟或定论客户历史记录！**

#### 2. Mandatory 3 Core Data Elements (必须明确的 3 大要素)
Whenever writing or populating the 6-month historical review narrative, the AI MUST obtain the following 3 verified data points:
1. **Total Samples Processed in Last 6 Months (分母：过去 6 个月送检总样本数)**:
   - Must be an exact number (e.g., `210 samples`, `660 samples`, `1,292 samples`).
   - Never write a vague phrase like `"processed samples"` without the exact numerical count.
2. **Prior Failing / Positive Occurrence Count (分子：历史不合格 / 阳性发生次数)**:
   - Must be verified: `0` (no prior occurrences) OR `N` occurrences ($N \ge 1$).
3. **Prior Incident Breakdown (如有历史阳性，必须提供完整清单)**:
   - If $N \ge 1$: Must provide OOS #, Sample ID (ETX), Submission Date, Analyte / Product Name, and Microbial Identification (Gram stain / morphology).
   - If $N > 3$: Must evaluate triggering Table 3 (Trend Table).

#### 3. Pre-Flight 2-Step Census Gate (前置必查与主动询问门禁)
At the start of drafting any OOS report, the AI must strictly execute this 2-step verification protocol:
- **Step 1 (查分子)**: 自动通过企业 SSO 访问 SharePoint 中央台账 `Sterile Lab - OOS Tracking Log.xlsx`，穿透检索对应的测试 Sheet（`Celsis Sterility OOS`、`Scan RDI OOS`、`<71> OOS`、`EM OOS` 等），提取确切先前 OOS 记录数及详细清单；
- **Step 2 (问分母)**: 若分母未在输入中提供，AI **必须主动向用户询问**该客户过去 6 个月的送检总数，获取核实后方可合流入报告，杜绝任何臆想！
  1. 过去 6 个月该客户在 Eagle 共送检了多少批次/样本进行该项检测？
  2. 过去 6 个月该客户/该产品是否有任何阳性或 OOS 历史记录？
  3. 若有历史阳性，具体的 OOS 编号、ETX 编号、日期及微生物鉴定结果是什么？

#### 4. Central Sterile Lab OOS Tracking Log Integration (SharePoint 中央 OOS 台账自动查验规范)
- **Authoritative Data Source**: The central live tracking workbook for the Sterile Microbiology Lab is hosted on SharePoint:
  `Sterile Lab - OOS Tracking Log.xlsx`
  Tabs include: `Celsis Sterility OOS`, `Scan RDI OOS`, `<71> OOS`, `EM OOS`, `USP <85> OOS`, `Particulate OOS`, `Media Fill OOS`, `OOS info for Priority Clients`, etc.
- **SSO Autonomous Access**: The AI can programmatically access SharePoint via Playwright using the persistent corporate SSO session (`LOCALAPPDATA/pastdue_playwright_session`) to export/download the latest workbook and run automated audits.
- **Role Differentiation (分子 vs 分母)**:
  - **Numerator (先前 OOS 记录数 $N$)**: Derived directly and strictly from the tracking log sheets by matching Client Name / Client ID / Method within the 6-month window ($T - 6\text{ months}$ to $T$).
  - **Denominator (送检总样本数 $Total$)**: Because the tracking log tracks OOS events, the overall sample count ($Total$) must be retrieved from LIMS / EagleTrax or confirmed with the user.

### 14. Phase 1 Standard Closing Sentence Rule (Phase 1 纯一阶段标准人类结语规范)
**CRITICAL**: When generating or reviewing Phase 1 OOS investigation reports:
1. **Classic Human Template Closing Only (只保留经典人类模板结语)**:
   The concluding paragraph (e.g., Page 5 `Text Field51` or Word report narrative conclusion) must strictly conclude with the standard, concise human sentence:
   `"Based on the laboratory investigation, no assignable laboratory cause was identified involving the analyst, instrument, reagents, supplies, or monitored laboratory environment. The initial [Test Method] OOS result therefore remains valid in accordance with the applicable laboratory OOS procedure."`
2. **Strict Ban on Redundant AI-Sounding Additions (严禁画蛇添足加长句)**:
   - **DO NOT** tack on redundant statements such as:
     `"Subsequent microbial identification testing did not recover an organism; consequently, the organism identity and source of the positive result could not be determined within the laboratory investigation."`
     (Microbial identification and subculture findings are already thoroughly documented in earlier sections/Paragraph 9; repeating them at the very end in a philosophical tone makes the report sound robotic and "unlike human writing").
   - **DO NOT** tack on:
     `"Further investigation and disposition, if required, should be performed in accordance with the applicable OOS procedure and client quality requirements."`
     (Standard workflow: Quality Assurance supervisors like Robin Seymour will review and route the document as standard procedure).
3. **Ergonomic & Natural Space Utilization**:
   Removing these redundant clauses ensures that Page 5 text flows naturally without overflowing the box, eliminating bottom voids and keeping the report 100% human-looking.

### 15. Human-Centric Typography, Punctuation & Method Designation Rule (极致拟人化标点与检测记录规范)
**CRITICAL**: Across all OOS investigation narratives (`ScanRDI`, `Celsis`, `USP <71>`, `Environmental Monitoring (EM)`):
1. **Elimination of Robotic Punctuation in Narrative Prose (杜绝机械符号，尽量用自然完整句子)**:
   - **Strict Ban on Semicolons (`;`)**: Human laboratory analysts rarely use semicolons in OOS narrative summaries. Split multi-clause ideas into separate, complete, and fluent sentences.
   - **Strict Ban on Dashes (`--` or `—`)**: Replace parenthetical dashes with natural prepositional phrases or appositives (e.g. change `"analysts – AC, GS – were interviewed"` to `"analysts involved in prepping and processing were interviewed, including Andrew Carrillo and Gabrielle Surber"`).
   - **Strict Ban on Colons (`:`) in Continuous Narrative**: Do not use colons like `Processing:` or `lot: LG342010349` in continuous narrative paragraphs. Use natural phrasing like `lot LG342010349` or `For Celsis processing, ...`.
   - **Minimal/Zero Parentheses in Narrative Prose**: Avoid excessive bracketed insertions like `(AC)`, `(114)`, `(01Sep2026)`, `(FTM cutoff = 2955 RLU)`, or `(Environmental Monitoring...)`. Embed these naturally into the grammatical flow of the sentence.
2. **Clarity on Method & Suitability Reference Numbers**:
   - Whenever citing internal method or suitability numbers (e.g. `2601120497`), always qualify with descriptive nouns such as `"suitability method record 2601120497"` or `"suitability test record 2601120497"` so any external auditor, client, or QA reviewer immediately understands what the number designates.
3. **Sentence-Driven Flow (尽量用完整句子)**:
   - Express thoughts in complete, professional, flowing sentences rather than fragmented clauses or bracketed shorthand.

### 16. Robin Seymour (QA Director) Sterility OOS Review Standards & Harmonization Rules (Robin 官方复核标准与合规八项铁律)
**CRITICAL**: Whenever drafting, reviewing, or compiling sterility OOS investigations (`Celsis`, `ScanRDI`, `USP <71>`):
1. **Point 1 — Investigation Procedure SOP Reference (首段强制引用 OOS 调查 SOP)**:
   - Paragraph 1 of Phase 1 narrative MUST explicitly include:
     `"This investigation was performed as per MICRO-SOP-53 Sterility Test Out-of-Specification (OOS) Investigation Procedure."`
   - Robin Seymour QA mandate: *"As there is an SOP for this, @Qiyue Chen @Olugbenga Ajayi let's start to make sure this information is added in future OOSs. Take note of this too."*
2. **Point 2 — Sample Storage Assessment (样本储存评价与人员资质严格剥离)**:
   - Page 1 `Text Field21` ("Was the sample stored appropriately?") and the narrative P2 must strictly state the sample was stored refrigerated (or specified temperature) per client instructions (e.g., `"Yes, the sample was stored refrigerated as per client's instructions"`).
   - Never conflate sample storage with analyst qualifications (`Text Field18`).
3. **Point 3 — Standard Used (Page 2 阳性对照与批号严禁重复堆叠)**:
   - In Page 2 `Text Field24`, write the reagent name: `Celsis ATP Positive Control`.
   - In `Text Field25` (Lot #), enter ONLY the lot number (e.g., `022601-1483`), NEVER repeating the reagent name in the lot section.
4. **Point 4 — Processing Cleanroom Suite Terminology (接种洁净室标准术语)**:
   - Standardize narrative description to: `"cleanroom suite used for processing procedures (CR115)"` (or `CR114` for Cleanroom Suite 114) rather than informal phrasing like `"cleanroom used for processing procedures (Suite 115)"`.
5. **Point 5 — Aliquoting Cleanroom Suite Terminology (分装洁净室标准术语)**:
   - Standardize narrative description to: `"cleanroom suite used for aliquoting procedures (CR114)"` rather than plain `"Suite 114"`.
   - Page 2 `Text Field32` (Equipment ID) must standardize cleanroom sensor designations to `E001736 (CR 114)` / `E001737 (CR 115)`.
6. **Point 6 — Total Floor Recovery Omission (彻底剔除分装与日常 EM 地面讨论)**:
   - Completely remove cleanroom floor recovery discussions from aliquoting and processing EM evaluations. Focus strictly on ISO 5 critical work surfaces, settling plates, operator glove touch plates, and active air monitoring.
7. **Point 7 — Weekly EM Assessment & Pre-Aliquoting Turbidity Rule (周检评价与分装前浑浊豁免准则)**:
   - If visible turbidity or microbial growth was already detected in the media bottle prior to the aliquoting step (e.g., recorded under an NCR), weekly EM for aliquoting is excluded from evaluation, with the explicit footnote/statement:
     `"Weekly environmental monitoring was not used in the evaluation of this OOS investigation given turbidity of media bottle from microbial growth was detected prior to aliquoting."`
   - If no turbidity was observed and the sample remained clear until Celsis analytical readout, weekly EM across the testing timeframe is evaluated, demonstrating that ISO 7 cleanrooms remained in control.
8. **Point 8 — Consolidated Environmental Monitoring Evaluation (环境监测统筹一体化防守闭环)**:
   - Consolidate processing and aliquoting EM discussions into a unified, airtight evaluation block:
     1. *Physical ISO 5 Containment*: Samples processed within validated ISO 5 BSCs during both processing and aliquoting.
     2. *Background Recovery Localization*: Any active air recovery in the outermost ISO 8 anteroom did not breach the ISO 7 buffer/cleanrooms or the ISO 5 critical zones.
     3. *Zero Pathway Proof*: 100% absence of microbial recovery across analyst glove plates, ISO 5 BSC surfaces, and settling plates proves that NO viable contamination transfer pathway existed.
     4. *Closed Transport*: Materials transported in disinfected, lidded bins on carts.
     5. *Analyst Compliance*: Strict adherence to MICRO-SOP-9 and MICRO-SOP-44 with zero deviations.

### 17. Balanced Dual-Page Visual Layout & Font Harmonization Standard (多页叙述段落均衡分配与字号完全对齐规范)
**CRITICAL**: Across all multi-page OOS AcroForm PDF generation:
1. **The Multi-Page AcroForm Dilemma**:
   - In PDF AcroForms, text boxes on separate pages (Text Field50 on Page 4 and Text Field51 on Page 5) cannot dynamically flow or bridge text across pages.
   - Cramming 11 paragraphs (~5,260 characters) into Page 5 while leaving Page 4 with only 5 paragraphs (~2,800 characters) forces Page 5 to shrink drastically to 7.325pt while Page 4 inflates to 12.0pt with an unsightly bottom void, creating an unhuman, jarring visual disparity.
2. **The 50/50 Balanced Character Distribution Mandate**:
   - Narrative paragraphs MUST be divided symmetrically by character volume (~4,500 characters on Page 4 and ~4,300 characters on Page 5).
   - **Page 4 (Testing Operations & Facility Baseline)**: Contains paragraphs p6..p12 (Sample receipt, direct inoculation in CR114B / BSC E001316, incubation & aliquoting in CR114A / BSC E001798, Celsis RLU readings, Subculture & Differential Staining 0 CFU, Media expiry & controls, Facility Monthly Cleaning of CR114 on 30 Aug 2026, and Introduction to Tables 2 & 3).
   - **Page 5 (Environmental Investigation & Quality Defense)**: Contains paragraphs p13..p21 (Processing EM review, Aliquoting EM review, Consolidated ISO 5 EM defense, Lack of contamination pathway, Analyst interview & cleaning compliance, Sample batch & cross-contamination review, Six-month client census, Lot history, and Final human conclusion).
3. **Harmonized Font Size (	ext_fontsize = 10.3pt)**:
   - Both Page 4 and Page 5 MUST be set to the identical, natural font size 10.3pt.
   - Both pages fill the available field box cleanly from top to bottom with ~27–29pt of natural bottom margin, completely eliminating bottom voids, preventing text truncation, and delivering a 100% natural, human-typed appearance.

### 18. Celsis Narrative Precision & 6-Month Analyte History Standard (Celsis 叙述精简与特定分析物历史规范)
**CRITICAL**: Across all Celsis Sterility OOS investigation reports:
1. **Omission of Negative TSB Numerical Cutoff Sentence (阴性管数值细节精简)**:
   - In Paragraph 8 (Celsis RLU reading summary), do NOT report the granular negative RLU and cutoff numbers for the negative TSB bottle (e.g. omit: `"The corresponding TSB sample container ETX-XXXXXX-XXXX tested negative with X RLU, below the TSB cutoff of X RLU and negative control of X RLU."`).
   - Simply state: `"All other FTM and TSB sample bottles in the same batch tested negative."`
2. **Omission of Lot Retesting Submission Sentence (批号复检历史段落剔除)**:
   - Do NOT include generic statements regarding whether the lot was submitted for retesting (e.g. omit: `"A review of the lot history shows that there have been no additional submissions of sample lot X for retesting..."`). Omit this paragraph entirely.
3. **Locked Standard Phrasing for 6-Month Analyte History (6个月特定分析物历史标准句式)**:
   - Standard template syntax:
     `Analyzing a 6-month sample history for [Client Name] indicates, this specific analyte "[Analyte / Sample Name]" has had no prior failures using the Celsis Sterility testing during this period.`
   - Concrete example:
     `Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "MOTs-C 10 MG/ML (5 ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.`


### 19. Strict 1-to-1 Transcription Protocol (转录任务纯搬运原则：纯复制粘贴，严禁改动字号与格式，顺其自然)
**CRITICAL**: When transcribing (转录) from an existing draft OOS document/PDF into the official ZenQMS investigation template (e.g. `CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf`):
1. **Pure Copy-Paste Only (纯复制粘贴，顺其自然，不做任何字号或留白干预)**:
   - Only copy `field_value` and checkbox states (`w.field_value`) 1-to-1 from the source document to the target document.
   - **DO NOT** attempt to calculate, override, or tamper with `text_fontsize` or text formatting.
   - **DO NOT** worry about bottom white space ("不要管留不留白，字体大小什么的，顺其自然就行了").
   - Whatever text is in the source field, copy and paste it directly. Let the PDF viewer and official template render the fields naturally.
2. **Zero Unauthorized Formatting Modifications**:
   - The user has explicitly mandated: "就转录（复制粘贴）过去就行，不要管什么字体大小格式的".
   - Treat the field contents as pure data transfer without adding manual typography styling or font overrides during transcription.

### 20. Strict Prohibition of Ampersand `&` & Informal Symbols (严禁使用 `&` 等非正式特殊符号铁律)
**CRITICAL**: In all cGMP formal reports, investigation forms (like `CORP-FORM-21`), comments, standalone tables, and narrative texts:
1. **Never Use `&` (Ampersand)**:
   - In formal regulatory documentation, `&` is considered an informal abbreviation that is strictly forbidden.
   - Always spell out `, and ` or ` and ` in full English sentences.
   - *Analyst Lists*: `"Yes, analysts Andrew Carrillo, Gabrielle Surber, America Alanis, and Cuong Du were comprehensively interviewed."` (NEVER `Alanis & Cuong Du`).
   - *Section Headers & Reagents*: `"Celsis Reagents and Kits:"` (NEVER `Reagents & Kits`).
2. **Zero Ampersand in Bacterial Tables**:
   - When listing multiple organisms recovered on a single plate in Table 2 or Table 3, **DO NOT join them with `&`**.
   - Separate the two organisms with a blank line between them.

### 21. Hyperlink Line Spacing & Vertical Baseline Alignment Invariance (超链接段落行距与水平基线绝对对齐铁律)
**CRITICAL**: Across all Word and PDF tables where external hyperlinks are injected:
1. **Absolute Ban on `line="360"` (严禁注入额外行距破坏垂直居中)**:
   - When injecting OpenXML `<w:hyperlink>` into table cells in Word/Docx, **NEVER inject `<w:spacing w:line="360" w:lineRule="auto"/>`**!
   - Extra line spacing expands the single-line bounding box and causes Word's vertical alignment (`w:vAlign="center"`) to center an inflated box, misaligning the hyperlinked Sample ID relative to adjacent cells (`1 CFU`, `None`).
2. **Preserve Native Cell Paragraph Properties**:
   - Keep standard paragraph spacing (no extra line spacing), inherit cell paragraph properties `<w:jc w:val="center"/>`, exact 7pt font (`w:sz="14"`), blue color (`#0000FF`), and single underline (`<w:u w:val="single"/>`).
   - In exported PDFs, the Sample ID hyperlinked text must match the exact vertical baseline (`y0`, `y1`) of adjacent cells down to 0.001pt precision.

### 22. Microbiological Binomial Nomenclature Formatting Standard (微生物拉丁双名法全斜体与隔行规范)
**CRITICAL**: In all OOS reports, investigation narratives, and tables:
1. **Mandatory Italics for Species (*Genus species*)**:
   - All bacterial genus and species names (*Corynebacterium ureicelerivorans*, *Mycobacterium grossiae*, *Micrococcus luteus*, *Staphylococcus epidermidis*, etc.) **MUST ALWAYS BE ITALICIZED** (`<w:i/>`, `<w:iCs/>` in Word; `<i>...</i>` or markdown `*...*` in docs).
   - Higher taxa (Family, Order) and informal descriptive terms (e.g. "Gram-positive rods", "Gram-positive cocci") remain in standard non-italic font.
2. **Multi-Organism Blank Line Separation**:
   - When multiple organisms are recovered from the same sampling site/plate (e.g., Colony 1 and Colony 2 on active air plate `ETX-260914-0487`), display each organism on its own line(s), separated by an empty blank line in between.
   - Do NOT concatenate with `&` or commas.

### 23. Sterility Sample Inoculation Volume Phrasing Standard (`per media type` 培养基接种体积术语铁律)
**CRITICAL**: In direct inoculation sterility testing (Celsis, USP <71>, Scan RDI), when the testing protocol divides the total sample volume across multiple containers/jars per media type (e.g., five 300 mL FTM jars and five 300 mL TSB jars):
1. **Mandatory Wording (`per media type`)**:
   - The narrative MUST strictly state:
     `"Direct inoculation was performed by adding [X] mL of sample per media type, using [N] [size] [Medium 1] jars and [N] [size] [Medium 2] jars."`
   - Example:
     `"Direct inoculation was performed by adding 25 mL of sample per media type, using five 300 mL FTM jars and five 300 mL TSB jars."`
2. **Strict Prohibition of `per media container`**:
   - **NEVER** write `"per media container"`.
   - Stating "per media container" erroneously implies that each individual jar received 25 mL (which would multiply the total sample volume tested to 250 mL, creating a serious factual contradiction against laboratory batch records and client submissions).

### 24. Suitability Method Record Notation & Positive Media Reading Conciseness Standard
1. **Suitability Record Notation (记录编号冒号规范)**:
   - When citing suitability method records in investigation narratives, always format with a colon after `record`: `"suitability method record: [RECORD_NUM]"` (e.g., `"suitability method record: 2601120497"`).
2. **Positive Container Redundancy Elimination (剔除阳性瓶子编号冗余后缀)**:
   - In the analytical reading summary sentence, do NOT append sub-container tracking IDs (e.g. `(ETX-260828-0527-3/5)`) after naming the positive media jar.
   - Standard sentence: `"sample ETX-XXXXXX-XXXX was found to yield a positive reading in one of the 300 mL FTM media jars."`
   - Rationale: Redundant sub-container IDs clutter narrative flow and are already explicitly and cleanly itemized in Table 1.

### 25. Sterility Positive Media Jar Count & Sequence Verification Gate (无菌检测阳性培养基瓶数与瓶序多重绝对校验铁律)
**CRITICAL & ZERO TOLERANCE**: In cGMP sterility investigations (Celsis, USP <71>, Scan RDI), misstating the count of positive media containers (e.g. reporting 1 positive jar when the raw data explicitly records 2 positive jars) is a catastrophic ALCOA+ data integrity failure that misrepresents the severity of contamination to QA and regulatory authorities.

1. **Raw Container Count Primacy (原始瓶数绝对守恒原则)**:
   - When raw laboratory notification emails or batch sheets state `$N \times \text{[size]} \text{ [media]}$` (e.g., `2 x 300mL FTM`), the integer $N$ is the immutable ground truth.
   - Table 1 Column `Media with microbial growth` MUST strictly match the exact count: `$N \times \text{[size]} \text{ [media]}$` (e.g., `2 x 300mL FTM`). NEVER reduce, abbreviate, or round down to `1 x`.

2. **Absolute Ban on Fraction Misinterpretation (严禁将多瓶序号压缩为分数导致单瓶误判)**:
   - Phrases indicating specific jar positions (e.g. `(3rd and 5th jars)` or `jars #3 and #5`) MUST NEVER be shorthand-compressed into notations like `3/5`.
   - The AI and scripts MUST NEVER interpret `3/5` as "bottle 3 of 5" (which caused the catastrophic error of reporting 1 jar instead of 2). Multi-jar enumerations must be explicitly preserved as distinct positive containers.

3. **Grammatical & Contextual Concordance Across All Narrative Sections (全篇前后文单复数严格一致性准则)**:
   - **When $N \ge 2$**:
     - Narrative Paragraph 8: MUST state `"sample [ETX] was found to yield positive readings in [word(N)] of the [size] [media] media jars ([explicit jars])"` (e.g., `"positive readings in two of the 300 mL FTM media jars (the 3rd and 5th jars)"`).
     - Referenced containers: MUST use plural nouns (`positive FTM sample bottles`, `duplicate reading tubes`, `bottles were submitted`).
   - **When $N = 1$**:
     - Narrative Paragraph 8: MUST state `"positive reading in one of the [size] [media] media jars"`.

4. **Pre-Delivery Raw Data Reconciliation Gate (交付前原始邮件与成品强制逐字穿透对账闸门)**:
   - Before declaring completion of ANY OOS report, the AI MUST execute an automated cross-reconciliation check:
     - Step 1: Scan user raw prompt/intake text for regex `\b(\d+)\s*x\s*\d+\s*mL\s*(FTM|TSB)`.
     - Step 2: Compare captured count $N$ against Table 1 cell and Narrative Paragraph 8 text.
     - Step 3: If any mismatch is detected ($N_{\text{raw}} \ne N_{\text{rendered}}$), execution MUST immediately abort and trigger self-correction before presenting results to the user.

### 26. Analyst Legal Full Name Precision & Two-Page Narrative Balance Standard (分析员真实法定全名与两页叙述平衡规范)
**CRITICAL**: In cGMP laboratory documentation and OOS reporting:
1. **Analyst Legal Name Accuracy (Mukyung Jang vs. Min Jang)**:
   - For initials MJ in Scan RDI sample preparation, the full legal system name is **Mukyung Jang (MJ)** (NEVER shorthand or truncated Min Jang).
   - Global Concordance: The name Mukyung Jang must be strictly harmonized across Section B comments (Text Field13), Page 3 narrative (Text Field49), and master Word report blocks.
2. **Two-Page Narrative Flow & Page 5 N/A Balance**:
   - When narrative content fits symmetrically across Page 3 (Text Field49, fs=8.75pt, ~3,540 chars) and Page 4 (Text Field50, fs=9.2pt, ~3,375 chars), Page 5 Text Field51 is marked as 'N/A QYC [Date]' to maintain clean visual balance and eliminate awkward bottom voids. Page 6 Text Field54 (Lab Manager) remains empty ('') for supervisor review.

### 27. Scan RDI Changeover (S/O) Bench Operator Precision & Clean EM Attribution Rule (Scan 换批操作员穿透与零微生物检出基线准则)
**CRITICAL**: In Scan RDI OOS investigations:
1. **S/O (Scan Changeover) Operator Primacy**:
   - The Changeover Analyst is the actual operator who performed the changeover procedure in the biological safety cabinet as signed on the physical bench logbooks.
   - For OOS-261967, bench records confirm Changeover was performed by **Karla Silva (KSM)** in BSC E001937, while Testing was performed by **Varsha Subramanian (VV)** in BSC E001319.
   - Page 1 Text Field3 must format all 4 roles distinctly:
     ```text
     Prepping Analyst: \rMukyung Jang (MJ)\r \rProcessing Analyst: \rVarsha Subramanian (VV)\r \rChangeover Analyst: \rKarla Silva (KSM)\r \rReading Analyst: \rSonal Uprety (SU)
     ```
   - Section B Interview Comment (Text Field13):
     `"Yes, analysts Mukyung Jang, Varsha Subramanian, Karla Silva, and Sonal Uprety were interviewed comprehensively."`
2. **Clean Environmental Monitoring Attribution**:
   - Both Testing (VV in BSC E001319) and Changeover (KSM in BSC E001937) personnel touch plates, surface contact plates, and settling plates demonstrated 100% absence of microbial recovery (No Growth / 0 CFU).
   - Table 2 explicitly captures separate rows for Testing and Changeover (S/O) across Personnel, Surface, and Settling sites, accurately reflecting their respective operators (VV and KSM) and BSC equipment IDs (BSC E001319 and BSC E001937).

### 28. Celsis Sterility Reviewer Standards & QA Golden Feedback (Robin Sharma / RS Review Feedback for OOS-262080)
**CRITICAL**: Across all Celsis Sterility OOS investigation reports, standalone tables, and narrative texts:
1. **Method Suitability Matching Negative Control Rule (适用性添加物负对照强制闭环准则)**:
   - When suitability method requires additives or neutralizers (e.g., BSA in PBS, peptone, polysorbate, etc.):
   - The narrative MUST explicitly state: `"Accordingly, a negative control with the same modifications was made."`
   - *Rationale*: Proves that the additive solution and modified media system remained sterile and did not introduce the contamination.
2. **Individual RLU Reporting for Multiple Positive Bottles (多瓶阳性独立 RLU 数值报告铁律)**:
   - When multiple containers/bottles test positive, do NOT report only a single combined average RLU across all bottles.
   - MUST report each positive bottle individually with its respective position and reading:
     `"The confirmed average Relative Luminescence Units (RLU) from the duplicate reading tubes, originating from the positive FTM sample bottles, yielded 4,202 RLU (Bottle 3) and 7,190 RLU (Bottle 5), both exceeding the positive cutoff of 2,955 RLU, where the FTM negative control was 985 RLU."`
3. **Colony Count Reconciliation in EM Active Air (活菌计数与菌种数量绝对统一准则)**:
   - When an active air sample recovers multiple bacterial isolates (e.g., *Corynebacterium ureicelerivorans* and *Mycobacterium grossiae* on `ETX-260914-0487`), the total colony count must accurately reflect all isolates (`2 CFU (ISO 8 114)`, NEVER understated as 1 CFU).
4. **Standard Subculture Negative Terminology ("Microbial growth could not be recovered")**:
   - In Table 1 Column `Microbial ID`, when subculture yields no growth for a positive sterility reading:
   - MUST write: `"Microbial growth could not be recovered"` (NEVER colloquial `"No growth was obtained"`).
5. **Universal EM Weekly Column Header ("Week of Testing")**:
   - In Tables 2 and 3, Column `Day /Week(s)` for weekly environmental monitoring rows MUST STRICTLY be:
     `"Week of Testing"` (NEVER `"Week on Testing Date"` or `"Week on Testing Date."`).

### 29. Anti-Robotic Human Narrative Standard & Zero Special Symbol Mandate (极致拟人化自然写作与全面清除机械符号永久铁律 - 🚨 终极合规规范)
**CRITICAL**: In all OOS reports, Word templates, AcroForm PDFs, comments, tables, and narrative texts:
1. **Zero Ampersand (`&`) Everywhere (绝对零 `&` 铁律)**:
   - NEVER use `&` as a conjunction in dynamic variables, analyst interview text, form fields, reagent boxes, or narrative prose.
   - Always write `, and ` or ` and ` in full English sentences.
   - *Example*: `America Alanis, and Cuong Du` (NEVER `Alanis & Cuong Du`); `Celsis Reagents and Kits` (NEVER `Reagents & Kits`).
2. **Complete Elimination of Bracketed Data Dumps (严禁括号式数据堆砌，全面转化为流利完整句子)**:
   - Real laboratory investigators and QA scientists never dump technical data into dense brackets `()`. Every piece of data must be woven into a fluent, grammatically complete sentence.
   - ❌ `positive readings in two of the 300 mL TSB media jars (Bottle 1 and Bottle 6)`
     ➡️ ✔️ `positive readings in two of the 300 mL TSB media jars, specifically Bottle 1 and Bottle 6`
   - ❌ `Bottle 1 yielded > 9,999,999 RLU (instrument overload)`
     ➡️ ✔️ `Bottle 1 yielded an instrument overload exceeding 9,999,999 RLU on both the initial read and confirmation re-read`
   - ❌ `Bottle 6 yielded 1,275 RLU (duplicate reading tubes 1,301 RLU and 1,249 RLU, %CV 2%)`
     ➡️ ✔️ `Bottle 6 yielded an average of 1,275 RLU with duplicate tube readings of 1,301 RLU and 1,249 RLU and a percent CV of 2 percent on the initial read`
   - ❌ `All other TSB bottles (Bottles 2, 3, 4, 5, 7, 8, 9)`
     ➡️ ✔️ `All other TSB bottles, including bottles 2, 3, 4, 5, 7, 8, and 9`
   - ❌ `(FTM cutoff 5,416.5 RLU, FTM negative control 1,806 RLU)`
     ➡️ ✔️ `where the FTM cutoff was 5,416.5 RLU and the negative control was 1,806 RLU`
   - ❌ `(< 30%)`
     ➡️ ✔️ `well within the acceptance criteria of less than 30 percent`
   - ❌ `Instrument Blank (8 RLU), Reagent Blank (77 RLU, %CV 1%), and ATP Positive Control (103,864 RLU, %CV 3%)`
     ➡️ ✔️ `including the instrument blank at 8 RLU, the reagent blank at 77 RLU with a 1 percent CV, and the ATP positive control at 103,864 RLU with a 3 percent CV`
   - ❌ `(TSB lot 07102026-1, Exp: 08Oct2026; FTM lot 06232026-5, Exp: 21Sep2026)`
     ➡️ ✔️ `with TSB lot 07102026-1 expiring on 08 Oct 2026 and FTM lot 06232026-5 expiring on 21 Sep 2026`
   - ❌ `on the date of testing (24Aug26), the preceding sampling date (21Aug26), or the subsequent sampling date (25Aug26)`
     ➡️ ✔️ `on the date of testing on 24 Aug 2026, the preceding sampling date on 21 Aug 2026, or the subsequent sampling date on 25 Aug 2026`
   - ❌ `recovery of 4 CFU (ETX-260901-0112) in the ISO 8 area, identified as 3 Gram (+) cocci and 1 Hyphae`
     ➡️ ✔️ `recovery of 4 CFU under test sample ETX-260901-0112 in the outermost ISO 8 anteroom, identified as three Gram-positive cocci and one Hyphae`
   - ❌ `cleanroom suite used for processing procedures (CR114)`
     ➡️ ✔️ `cleanroom suite used for processing procedures, cleanroom suite 114`
   - ❌ `(OOS-261500 and OOS-261878)`
     ➡️ ✔️ `documented under OOS-261500 and OOS-261878`
   - ❌ `(Tesamorelin 12 mg per Vial)`
     ➡️ ✔️ `for Tesamorelin 12 mg per Vial`
3. **Zero Robotic Colons in Continuous Prose (段落中彻底剔除伪小标题冒号)**:
   - 严禁在连贯叙述中插入像 `Evaluation of Environmental Monitoring Results:` 这样的突兀冒号。
   - 必须用自然的引导从句过渡：`Regarding the evaluation of environmental monitoring results, it is important to note that...`。
   - 引用适用性记录编号时，严禁加冒号：`suitability method record ETX-251218-0432`（严禁 `suitability method record: ...`）。

### 30. Scientific Units Preservation & Dual Date Format Standard (DDMMMYY 与 MMM YYYY 规范及科学专业单位守恒铁律 - 🚨 终极标准)
**CRITICAL**: Across all OOS investigation reports, standalone tables, and narrative texts:
1. **Professional Scientific Units Preservation (专业科学符号与单位绝对保留，严禁误伤硬改成英文单词)**:
   - 专业符号与度量衡单位（如 `%`, `%CV`, `CV%`, `℃`, `CFU`, `CFUs`, `RLU`）属于 cGMP 实验室的标准学术表达，**绝对不能当作“特殊符号”被误伤硬改成英文单词**！
   - `%` / `%CV` / `CV%`：必须保留符号（如 `2%`, `%CV < 30%`, `CV% of 3%`，严禁改成 `percent` 或 `percent CV`）。
   - `℃`：保留摄氏度单位符号（如 `30 to 35 ℃`, `20 to 25 ℃`，严禁写成 `degrees Celsius` 或 `degrees C`）。
   - `CFU` / `CFUs`：保留微生物菌落计数单位。
   - `RLU`：保留荧光读数单位。
   - “少用特殊符号”仅针对非正式文本符号（如 `&` 代替 `and`、段落中滥用伪小标题冒号 `:`、或把主谓宾数据丢进密集括号 `(...)`）。
2. **Dual Date Format Standard in Narrative Prose (正文叙述日期双轨制铁律)**:
   - **具体日期（有 DD）**：**必须使用 `DDMMMYY` 紧凑格式**（如 `21Aug26`, `24Aug26`, `31Aug26`, `12Jul26`, `16Aug26`, `25Aug26`, `28Aug26`, `01Sep26`, `04Sep26`）。中间严禁带空格，严禁写成 `24 Aug 2026`。
   - **宽泛日期（无 DD，仅年月）**：**必须使用 `MMM YYYY` 格式**（如 `Jul 2026`, `Aug 2026`, `Sep 2026`, `Jan 2027`）。
3. **Pruning & High-Level Narrative Standard (去粗取精与正文高级叙述原则)**:
   - 耗材批号/效期（如 Tween 80 批号、TSB/FTM 批号）已在 Page 2 耗材栏详细列出，正文段落无需重复罗列其批号与效期。
   - 仪器设备（如 Celsis 读数仪）正文中直接写作 `Celsis E002222`，无需长串修饰 `Advance 2, equipment ID E002222`。
   - 历史记录（6个月记录）直奔主题，直接指明该特定分析物（Tesamorelin 12 mg per Vial）的历史阳性次数与 OOS 编号（OOS-261500 和 OOS-261878），剔除无意义的批号复测套话。
   - 同批交叉污染段落精炼为明确的 3 句话，无需冗余罗列 6 个同批阴性 ETX 编号。
4. **Logical 3-Page Flow (三页黄金结构)**:
   - Page 3 (`Text Field49`): 实验前准备、洁净室、接种与分装（P1..P7）
   - Page 4 (`Text Field50`): Celsis 读数结果、菌种确认、效期、月度清洁与表格引言（P8..P12）
   - Page 5 (`Text Field51`): 环境监测调查、交叉污染排除、6个月历史审查与最终结论（P13..P21）

### 31. Universal Mandatory Blank Line Separation for Multi-Item Table Cells (表格多菌种/多记录单元格绝对强制空行隔开永久铁律 - 🚨 终极防线)
**CRITICAL**: Across all OOS modules (`Scan RDI`, `Celsis`, `USP <71>`, `Environmental Monitoring (EM)`), and across all table formats (Table 1, Table 2, Table 3, and Standalone Tables):
1. **Zero Tolerance for Adjacent Unseparated Items (相邻条目未空行零容忍准则)**:
   - When a table cell contains multiple items (such as multiple microbial isolates/species, multiple CFU counts, or multiple ETX plate numbers), **EVERY SINGLE ITEM MUST BE SEPARATED BY A DEDICATED EMPTY BLANK LINE (`\n\n` in text, or a blank spacing paragraph `<w:p>` in Word OpenXML)**.
   - **NEVER group or place any two items together on adjacent lines without a blank line in between!**
2. **Distinct Microbial Strains Separation (菌种鉴定栏目绝对空行)**:
   - In `Microbial ID` / `Related Microbial ID` cells, every distinct isolate MUST be separated from adjacent isolates by a clear empty line:
     *Kocuria indica*
     [空行]
     *Brevibacterium sp. CS2*
     [空行]
     *Corynebacterium sp*
     [空行]
     *Micrococcus luteus*
     [空行]
     *Paracoccus yeei*
     [空行]
     *Staphylococcus hominis*
     [空行]
     *Kocuria rhizophila*
   - Strictly forbidden: Placing *Corynebacterium sp* and *Micrococcus luteus* on adjacent lines without a blank line.
3. **Multi-Hit EM Observations & ETX Numbers (多点检出环境监测栏目横向对齐空行)**:
   - In `Observation` cells with multiple hits (e.g. L-Suite active air):
     `6 CFUs (ISO 8 143 Sec I)`
     [空行]
     `8 CFUs (ISO 8 143 Sec II)`
     [空行]
     `2 CFUs (ISO 8 142)`
   - In `EM Plate ETX Number` cells, corresponding ETX plate numbers MUST also be separated by matching empty lines to maintain strict horizontal alignment:
     `ETX-260929-0335`
     [空行]
     `ETX-260929-0341`
     [空行]
     `ETX-260929-0344`
4. **Implementation & Pre-Delivery Audit Gate (底层代码与交付前自检闸门)**:
   - In Python scripts, always join multi-item lists with `\n\n`, never `\n`.
   - In Word OpenXML, append explicit empty spacing paragraphs `<w:p>` with proper line spacing.
   - Always run pre-delivery cell line-spacing verification before finalizing documents.

### 32. Sterility Investigation Phase 1 vs. Phase 2 Scope Separation Architecture (无菌检测 Phase 1 与 Phase 2 调查范围绝对解耦架构 - 🚨 黄金标杆规范)
**CRITICAL**: When an OOS sterility investigation involves both an initial test (Original Test) and a retest (Retest) (Reference: `OOS-261987 Uriel Pharmacy P1 & P2 Signed by OA-RS`):
1. **Master Table Document (`Tables ... DOCX/PDF`)**:
   - Contains BOTH tests: Table 1 lists both original and retest samples; Table 2 lists original EM; Table 3 lists retest EM. The table contract preserves full transparency across both stages.
2. **Phase 1 Official Investigation Report (`CORP-FORM-21 / 3.100.019.F01`)**:
   - **MUST STRICTLY INVESTIGATE THE ORIGINAL TEST ONLY!**
   - Page 1: Sample ID is ONLY the original test (e.g. `ETX-260902-0505`), test date is ONLY the original test date (`04-Sep-2026`), analysts are ONLY the original analysts (`Alex Saravia` / `Elysse Nioupin`), description of incident covers only initial failure.
   - Page 2: Equipment is ONLY the original BSC (`ISO 5 BSC E001316`) and Cleanroom (`CR114`).
   - Page 3..Page 5: Narrative focuses strictly on the initial test operations, reading, preliminary organism identification, initial EM bracketing defense, and ruling out Phase 1 lab error. **NEVER mention retest, retest analyst, or retest sample ID in Phase 1!**
   - Page 6: Concludes initial test is valid (`Check Box88` = Yes).
3. **Phase 2 Official Investigation Report (`CORP-FORM-22 / 3.100.019.F02`)**:
   - Dedicated to the **Retest Investigation**.
   - Page 1: Records both Original and Retest IDs, Retest Date, Retest Analyst, and Retest Result.
   - Page 3: Recaps the initial test investigation.
   - Page 4: Detailed retest narrative under `RETEST UNDER SUBMISSION ETX-XXXXXXXX-XXXX`, comparing kinetics and organism sequencing.
   - Page 5: Identifies root cause (Product Contamination / External Phenomena) and closes investigation.





