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
    *   **2026-09-16 Update 1:** Integrated latest RS feedback (OOS-261878) to explicitly separate and track the **Reading Analyst** (adding them to the interview statement and deduplicated analyst list), added explicit "viable but not culturable" terminology when no growth is recovered during Differential Staining, and finalized the conclusion statement to strictly state "laboratory error is highly unlikely" without recommending additional evaluation.
    *   **2026-09-16 Update 2:** Integrated RS feedback (OOS-261877) to remove "CR" and "Suite" prefixes from cleanroom references (e.g., "114", "114A", "114B"), change formula worksheet verification to "N/A", and condense the Smart Justification Engine's defense for floor recoveries to explicitly state that the recovery was from a non-critical area physically separated from the ISO 5 processing zone with no viable transfer pathway.
*   **Smart Justification Verbiage:** The precise sentences used in the `smart_just` block (e.g., "Notably, the colony morphology...", "It is important to note that the microbial recovery occurred in a non-critical area...") are finalized and locked based on user approval.
*   **Cross-Contamination & Lot History Logic:** Dynamic generation of text regarding sample-to-sample contamination and lot history checking is finalized.
*   **EM Module 5-Template Ecosystem & 7/8-Page PDF Pipeline:** EM OOS module templates (`EM OOS P1 template.docx`, `EM OOS P1 template 0.docx`, `EM OOS P1 template.pdf`, `tables for em.docx`, dynamic Page 7/8 Table PDF generator), `em_logic.py`, `pages/EM.py`, and `create_new_module.py` are live and fully verified against production EM reports (including `OOS-261187` and `OOS-261186`).
    *   **2026-09-21 Update (OOS-261186 Review & Package Assembly):** Established complete 8-page package standard (6-page ZenQMS form + Page 7 Version History + Page 8 Table 1 & Table 2 attachment). Verified interactive AcroForm text field updates (`Text Field11` font 6.5pt for `Action level: >= 10 CFU/Plate`, `MICRO-SOP-2 Rev 16 (23-Jul-2026)`, Suite 116 correction, and preservation of Eagle legacy EM transient root cause boilerplate `may be attributed to a potential analyst error` per laboratory precedent). Deliverables strictly standardized as `<OOS-ID> <Sample> - EM.pdf/.docx` and `Tables <OOS-ID> <Sample> - EM.pdf/.docx`.
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
    8. *Scan RDI Table 2 Single-BSC Deduplication, No Date in BSC Header & Notes Column Widening:* In `tables for scan.docx`, `ScanRDI OOS P1 template.docx`, `pages/ScanRDI.py`, and `scratch/process_oos_261814.py`: (a) When testing and changeover BSC are identical (`bsc_id == chgbsc_id`), the BSC bracketing header dynamically formats to single BSC without date (`Biological Safety Cabinet EM Bracketing Biological Safety Cabinet (BSC) E00{bsc_id}`), and duplicate changeover Surface and Settling rows are automatically removed. (b) All EM Notes (e.g. `plate was desiccated on 5 day read.`) are dynamically mapped via Jinja2 variables (`{{ note_pers }}`, `{{ note_surf }}`, `{{ note_sett }}`, `{{ note_air }}`, `{{ note_room }}`) into Table 2/3 instead of hardcoded `'None'`. (c) The Notes column width is expanded from 390,525 to 1,014,090 EMUs (~1.11 inches, >2.6x wider) across all table rows and `tblGrid`, preventing crowded text wrapping.
    9. *Trend Table (Table 3) Dynamic Thresholding:* Table 3 ("In Trend of Past OOS Results...") and its associated header paragraph are strictly conditional: they are ONLY included when there are more than 3 past OOS records (`prior_count > 3`). If 3 or fewer records exist (or 0 prior failures), the table and paragraph are automatically excised from generated documents.
    10. *Universal Table 1 Clickable Hyperlink Standard & Word-to-PDF Conversion (Table 1 原生交互超链接常态化规范):* Across ALL OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `EM`), whenever Table 1 is generated (in standalone table documents or master report Page 7/8 attachments), the Sample ID (`ETX-XXXXXX-XXXX`) **MUST ALWAYS** be embedded with an active, clickable external hyperlink directly pointing to its EagleTrax test details URL (`https://etrax.eagleanalytical.com/SubmissionTest/Details/...` or `/Submission/Details/...`).
        - *Technical Implementation:* The link is constructed via OpenXML `<w:hyperlink>` with `w:rStyle="Hyperlink"` and standard `#0000FF` blue underline formatting in `python-docx`.
        - *PDF Interactive Preservation:* When converted via Word COM, the interactive link is natively compiled into the PDF Annotation tree, ensuring QA reviewers (Robin/Neha) and management can open the live EagleTrax sample page with a single click directly from any PDF viewer.
        - *Mandatory Execution:* This is NOT optional. Whenever an ETX URL or submission link is provided or resolved, it MUST be attached to the Sample ID in Table 1 across both Word and PDF deliverables.
    11. *Same-Day Downloaded QA Form (CORP-FORM-21) Transcription & 7-Page PDF Integration Pipeline:* Per QA compliance requirement that investigation forms must be freshly downloaded from ZenQMS on the day of investigation: (a) The generator transcribes the complete OOS dataset directly into the downloaded `CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf`, populating all 65 standard compliance checkboxes and boilerplate alongside dynamic sample fields. (b) Table 2 Notes strictly enforce title capitalization and no trailing period (`Plate was desiccated on 5 day read`). (c) The standalone Table PDF is appended as Page 7 to the 6-page form, producing a unified 7-page official OOS PDF report.
    *   **2026-09-25 Update (OOS-262080 Full Celsis Production Package & Ground Truth Audit):** Finalized the production pipeline for Celsis OOS-262080 (ETX-260828-0527, Optimal Balance Pharmacy). Enforced 5 cGMP investigation rules: (1) Spatial Topology verification (BSC 1316 in 114B / Suite 114 strictly requires Suite 114 weekly EM; rejecting Suite 115 trap); (2) Dual-Track EM Tracing (Personnel EM follows CCD personal shifts on 03Sep/08Sep/15Sep, Settling & Surface follow BSC 1798 hood history on 07Sep ALA/08Sep CCD/09Sep ALA); (3) Raw Bench Records Precedence (Cuong Du / CCD designated as Aliquoting Analyst based on physical logbook signatures); (4) Dual-Week ISO 8 Anteroom 1 CFU defense integration (04Sep26 ISS under ETX-260914-0487 and 10Sep26 SMO under ETX-260921-0520 defended via physical segregation and 0 CFU critical zone recovery); (5) Post-processing defensive cell overwrite for hardcoded template text. All 5 deliverables generated and verified on Desktop.
    12. *Celsis OOS Full Pipeline & Table Bracketing Standard (OOS-262080):* Verified against production OOS-262080 (ETX-260828-0527): (a) Strict 4-analyst tracking across Prepper, Processor, Aliquoter, and Reader with explicit interview statements. (b) Suite cleanrooms referenced strictly without prefixes ("114", "114A", "114B"), Formula Worksheet marked "N/A" with initials/date. (c) Condensed Smart Justification defending non-critical floor/air recoveries (e.g. 1 CFU active air in ISO 8 114) by physical segregation and transfer pathway cutoff (all ISO 5 BSC surfaces, settling, and analyst gloves = 0 CFU / No Growth). (d) Two-phase Table 2 (Processing) and Table 3 (Aliquoting) formatting with automated PageBreak before Table 3 to eliminate orphaned headers. (e) Produces 5 synchronized deliverables: Main DOCX, Standalone Tables DOCX, Main PDF, Standalone Tables PDF, and Complete Combined PDF package.
    *   **2026-09-29 Update (OOS-261967 Scan RDI Package & EM Logbook Census Rule):** Finalized the production pipeline for Scan RDI OOS-261967 (ETX-260813-0778, TAM Pharmacy). Enforced the Mandatory EM Logbook Census Protocol and incorporated the 6-page comprehensive EM scan.
    13. *Mandatory EM Logbook Census Protocol & Anti-Extrapolation Rule (EM 台账严谨清点与零外推防陷阱法则):*
        *   **The 6-Document Complete EM Census Checklist (台账6大要素完整性清点清单):** Whenever an environmental monitoring scan PDF (or image set) is provided, the AI MUST systematically inspect the Document Number in the header block of every page and verify the presence of all 6 standard microbiology logbooks before declaring any EM investigation dataset complete:
            1. `3.600.002.F01` (Settling Sampling Logbook): Passive microbial settling in ISO 5 Critical Zone (BSC/LAFW).
            2. `3.600.002.F02` (Surface Sampling Logbook): Contact plates 1–4 on ISO 5 BSC interior surfaces.
            3. `3.600.002.F03` (Personnel Sampling Logbook): Operator left and right fingertip touch plates.
            4. `3.600.002.F04` (Anteroom/Buffer Room Surface Sampling Logbook): ISO 7 buffer and ISO 8 anteroom contact plates.
            5. `3.600.002.F05` (Active Air Sampling Logbook): Cleanroom volumetric active air monitoring plates.
            6. `3.600.018.F05` (BSC Cleaning & Disinfection Logbook): Daily disinfection and contact time verification.
        *   **Strict Anti-Extrapolation Principle (零外推、零臆想铁律：F02 ≠ F01):** Having `F02` (Surface) does NOT mean `F01` (Settling) exists or is negative. They are completely separate physical logbooks with separate SOP forms, incubation tracks, and physical plates. If any logbook form (specifically `3.600.002.F01` Settling) is absent from the provided PDF scan, the AI MUST NOT assume "No Growth" or extrapolate from other plates. It MUST explicitly flag the exact missing document number and title to the user, and mark that specific sampling site as `[Pending]` in all draft tables and narratives until the physical scan is supplied.
        *   **Pre-Response Document Verification Gate (在回答“还缺什么”前的必检门禁):** When the user asks "还缺什么吗？" (Is anything missing?) or presents raw EM scans, the AI is STRICTLY FORBIDDEN from answering casually. The AI MUST systematically inspect each page of the PDF, extract the Document # (`DOCUMENT #: 3.600.002.F0X`) and Title in the header block, map them against the 6-document checklist, and report: (a) Form identified per page; (b) Target Date, Analyst, and BSC match; (c) Any missing form(s) explicitly listed by Document # and Name.
        *   **Multi-Page Rotation & Full-Page Verification:** Scanned PDFs from multifunction copiers frequently rotate pages 90°, 180°, or 270°. The AI must systematically check or normalize page orientation to read every header and footnote, ensuring no page is misread due to inversion.
    14. *Automated EagleTrax Direct Scrape Protocol & EM Plate Hyperlink Standard (EagleTrax 自动化查询与 EM 超链接规范):*
        *   **Automated Corporate SSO / Session Re-use:** When provided with EagleTrax test links or ETX numbers, the AI can programmatically access EagleTrax via Playwright utilizing persistent corporate SSO session cookies (`LOCALAPPDATA/pastdue_playwright_session`).
        *   **Automated Morphology Extraction:** Directly extract live Differential Staining Gram stain observations (`Gram (+) rods`, `Gram (+) cocci`, etc.) from `#TestDetails` and capture verification screenshot evidence, eliminating manual copy-pasting.
        *   **EM Plate ETX Interactive Hyperlinks:** In Table 2, all positive EM Plate ETX numbers (e.g. `ETX-260908-0584` and `ETX-260908-0580`) **MUST ALWAYS** be embedded with native OpenXML clickable hyperlinks pointing directly to their live test details URL, matching the Table 1 Sample ID hyperlink standard.
        *   **Windows File-Lock Defense:** When generating Word and PDF deliverables, scripts must compile through temporary files (`temp_tables_export.pdf`) to gracefully handle file locks when Acrobat or Edge has the target desktop PDF open.
    15. *Natural Human Typography & Zero Artificial Padding Standard (原生手打字号与杜绝人为缩放留白常态化规范):*
        *   **Strict Prohibition of Manual `text_fontsize` Overrides:** Across ALL Eagle PDF templates (`ScanRDI`, `Celsis`, `USP <71>`, `EM`), scripts and agents are strictly forbidden from manually setting fractional or shrunken font sizes (e.g. `w.text_fontsize = 5.8`, `6.0`, `6.2`, `6.5`, `7.5`, `8.5`). Text fields must retain their native template Auto/default appearance (`size = 0.0` / `/DA`), mirroring real human input in Adobe Acrobat (~9-10pt).
        *   **Zero Artificial Padding:** Under no circumstances may scripts append trailing carriage returns or blank lines (`\r \r \r \r`) to create artificial bottom spacing or shifts. Text must end cleanly and naturally.
        *   **Incubator Spacing Invariance:** In Page 2 `Text Field43` and `Text Field44`, sensor names remain on the same line as the incubator (`Incubator E001356 (Sensor E001450)`), each incubator entry is separated by a single empty line (`\r \r`), aligning 1-to-1 with calibration dates (`Jan 2027 / Feb 2027\r \r`).
        *   **Universal Applicability:** Permanently locked for all current and future OOS automation workflows.
    16. *Cleanroom Facility Monthly Cleaning Schedule & Prior Bracketing Standard (MICRO-SOP-9):*
        *   **Bi-Weekly Sunday Cleaning Cadence:** Facility monthly cleaning and disinfection (utilizing $H_2O_2$ vapor / chemical indicators across cleanroom suites and BSCs) operates on an official dual-Sunday schedule per month: **Mid-Month Sunday** and **Last Sunday**.
        *   **Authoritative 2026 Facility Schedule:**
            - August 2026: August 16, 2026 & August 30, 2026
            - September 2026: September 13, 2026 & September 27, 2026
            - October 2026: October 11, 2026 & October 25, 2026
            - November 2026: November 15, 2026 & November 29, 2026
            - December 2026: December 13, 2026 & December 27, 2026
        *   **Prior Bracketing Logic & Anti-Post-Event Invariance:** When generating investigation narratives, the cleaning date is strictly resolved to the most recent cleaning event occurring **BEFORE** the sample processing/testing date ($\text{cleaning\_date} < \text{event\_date}$). **Strictly forbids citing cleaning dates that occurred AFTER processing** (no forward-looking post-hoc validation; eliminates month-name cognitive trap).
    17. *Mandatory 6-Month Client Sample History Census Protocol (客户 6 个月历史记录严谨审查与防臆想门禁规范 - 🚨 重大错误警示与永久铁律):*
        *   **🚨 血的教训与致命错误警示 (Lesson Learned & Permanent Ban):** 曾发生重大失误（起草 OOS 报告时未向用户核实客户 6 个月历史送检记录，或使用无真实数据支撑的套话 "processed samples with no prior occurrences"）。在 cGMP / FDA 审计标准下，任何未经验证的历史陈述均属严重合规漏洞！现已永久固化进 PRD：**严禁在未穿透核实真实数据前私自草拟或定论客户历史记录！**
        *   **3 Core Required Elements (必须获取的 3 大要素):**
            1. 分母：过去 6 个月送检总样本数（Total Samples Processed，精确数字，如 `120 samples`，严禁写模糊的 "samples"）；
            2. 分子：先前 OOS / 阳性发生次数（Prior Occurrences Count，0 或 N）；
            3. 清单：若 $N \ge 1$，必须附带完整 OOS #、ETX #、检测日期、产品名称与微生物鉴定结果（若 $N > 3$ 必须触发趋势表 Table 3）。
        *   **Pre-Flight Proactive Prompt Gate (前置必查与主动询问门禁):** 调查启动前，AI 必须主动执行两步核验：
            1. **Step 1 (查分子)**：通过企业 SSO 自动访问并审计 SharePoint 中央台账 `Sterile Lab - OOS Tracking Log.xlsx`，穿透检索对应的测试 Sheet（`Celsis Sterility OOS`、`Scan RDI OOS`、`<71> OOS`、`EM OOS` 等），提取确切先前 OOS 记录数及详细清单；
            2. **Step 2 (问分母)**：若分母未提供，AI 必须主动向用户询问该客户过去 6 个月的送检总数，获取核实后方可合流入报告，杜绝任何臆想！
    18. *Central Sterile Lab OOS Tracking Log Protocol (SharePoint 中央 OOS 台账自动查验规范):*
        *   **Authoritative SharePoint Source:** The central tracking workbook for the Sterile Microbiology Lab is hosted on SharePoint: `Sterile Lab - OOS Tracking Log.xlsx`. Tabs include: `Celsis Sterility OOS`, `Scan RDI OOS`, `<71> OOS`, `EM OOS`, `USP <85> OOS`, `Particulate OOS`, `Media Fill OOS`, `OOS info for Priority Clients`, etc.
        *   **SSO Autonomous Access:** The AI programmatically accesses SharePoint via Playwright using persistent corporate SSO session (`LOCALAPPDATA/pastdue_playwright_session`) to export/download the latest workbook and run automated multi-sheet audits.
        *   **Numerator vs. Denominator Protocol:** Prior OOS count ($N$) is extracted directly and verified against the log. Denominator (total samples processed in 6 months) is obtained from EagleTrax/LIMS or confirmed with the user.
    19. *Phase 1 Standard Closing Sentence Rule (Phase 1 纯一阶段标准人类结语规范):*
        *   **Classic Human Template Closing Only:** For Phase 1 investigations, the concluding narrative paragraph must strictly end with the concise, authentic human template sentence:
            `"Based on the laboratory investigation, no assignable laboratory cause was identified involving the analyst, instrument, reagents, supplies, or monitored laboratory environment. The initial [Test Method] OOS result therefore remains valid in accordance with the applicable laboratory OOS procedure."`
        *   **Strict Ban on Redundant AI-Sounding Additions (严禁画蛇添足加长句):**
            - **DO NOT** tack on: `"Subsequent microbial identification testing did not recover an organism; consequently, the organism identity and source of the positive result could not be determined within the laboratory investigation."` (already stated in the subculture section; adding it here sounds robotic and "unlike human writing").
            - **DO NOT** tack on: `"Further investigation and disposition, if required, should be performed in accordance with the applicable OOS procedure and client quality requirements."` (supervisors like Robin Seymour perform routine review and disposition routing).
            - Keep the conclusion natural, concise, and 100% human-like.
    20. *Human-Centric Typography, Punctuation & Method Designation Rule (极致拟人化标点与检测记录规范):*
        *   **Elimination of Robotic Punctuation in Narrative Prose (杜绝机械符号，尽量用自然完整句子):**
            - **Strict Ban on Semicolons (`;`):** Human laboratory analysts rarely use semicolons in OOS narrative summaries. Split ideas into separate, complete, and fluent sentences.
            - **Strict Ban on Dashes (`--` or `—`):** Replace parenthetical dashes with natural prepositional phrases or appositives (e.g. change `"analysts – AC, GS – were interviewed"` to `"analysts involved in prepping and processing were interviewed, including Andrew Carrillo and Gabrielle Surber"`).
            - **Strict Ban on Colons (`:`) in Continuous Narrative:** Do not use colons like `Processing:` or `lot: LG342010349` in continuous narrative paragraphs. Use natural phrasing like `lot LG342010349` or `For Celsis processing, ...`.
            - **Minimal/Zero Parentheses in Narrative Prose:** Avoid excessive bracketed insertions like `(AC)`, `(114)`, `(01Sep2026)`, `(FTM cutoff = 2955 RLU)`, or `(Environmental Monitoring...)`. Embed these naturally into the grammatical flow of the sentence.
        *   **Clarity on Method & Suitability Reference Numbers:** Whenever citing internal method or suitability numbers (e.g. `2601120497`), always qualify with descriptive nouns such as `"suitability method record 2601120497"` or `"suitability test record 2601120497"` so any external auditor, client, or QA reviewer immediately understands the document context.
    21. *Robin Seymour (QA Director) Sterility OOS Review Standards & Harmonization Rules (Robin 官方复核标准与合规八项铁律):*
        *   **Point 1 — Investigation Procedure SOP Citation (首段强制引用 OOS 调查 SOP):**
            - The very first narrative paragraph (Page 3 `Text Field49` / Word report P1) MUST explicitly conclude with:
              `"This investigation was performed as per MICRO-SOP-53 Sterility Test Out-of-Specification (OOS) Investigation Procedure."`
            - Robin Seymour explicit QA directive: *"As there is an SOP for this, @Qiyue Chen @Olugbenga Ajayi let's start to make sure this information is added in future OOSs. Take note of this too."*
        *   **Point 2 — Sample Storage Assessment (样本储存与资质严格剥离):**
            - Page 1 `Text Field21` ("Was the sample stored appropriately?") and the narrative P2 must strictly state the sample was stored refrigerated (or at requested condition) per client instructions (e.g., `"Yes, the sample was stored refrigerated as per client's instructions"`).
            - Strictly prevent conflating sample storage with analyst training or qualification (`Text Field18`).
        *   **Point 3 — Standard Used (Page 2 阳性对照与批号严禁重复堆叠):**
            - In Page 2 `Text Field24`, write the reagent name: `Celsis ATP Positive Control`.
            - In `Text Field25` (Lot #), enter ONLY the lot number (e.g., `022601-1483`), NEVER repeating the reagent name in the lot section.
        *   **Point 4 — Processing Cleanroom Suite Terminology (接种洁净室标准术语):**
            - Standardize the narrative description to: `"cleanroom suite used for processing procedures (CR115)"` (or `CR114` for Cleanroom Suite 114) rather than informal phrasing like `"cleanroom used for processing procedures (Suite 115)"`.
        *   **Point 5 — Aliquoting Cleanroom Suite Terminology (分装洁净室标准术语):**
            - Standardize the narrative description to: `"cleanroom suite used for aliquoting procedures (CR114)"` rather than plain `"Suite 114"`.
            - Page 2 `Text Field32` (Equipment ID) must standardize cleanroom sensor designations to `E001736 (CR 114)` / `E001737 (CR 115)`.
        *   **Point 6 — Total Floor Recovery Omission (彻底剔除分装与日常 EM 地面讨论):**
            - Completely remove cleanroom floor recovery discussions from aliquoting and processing EM evaluations. Focus strictly on ISO 5 critical work surfaces, settling plates, operator glove touch plates, and active air monitoring.
        *   **Point 7 — Weekly EM Assessment & Pre-Aliquoting Turbidity Rule (周检评价与分装前浑浊豁免准则):**
            - If visible turbidity or microbial growth was already detected in the media bottle prior to the aliquoting step (e.g., recorded under an NCR), weekly EM for aliquoting is excluded from evaluation, with the explicit footnote/statement:
              `"Weekly environmental monitoring was not used in the evaluation of this OOS investigation given turbidity of media bottle from microbial growth was detected prior to aliquoting."`
            - If no turbidity was observed and the sample remained clear until Celsis analytical readout, weekly EM across the testing timeframe is evaluated, demonstrating that ISO 7 cleanrooms remained in control.
        *   **Point 8 — Consolidated Environmental Monitoring Evaluation (环境监测统筹一体化防守闭环):**
            - Consolidate processing and aliquoting EM discussions into a unified, airtight evaluation block:
              1. *Physical ISO 5 Containment*: Samples processed within validated ISO 5 BSCs during both processing and aliquoting.
              2. *Background Recovery Localization*: Any active air recovery in the outermost ISO 8 anteroom did not breach the ISO 7 buffer/cleanrooms or the ISO 5 critical zones.
              3. *Zero Pathway Proof*: 100% absence of microbial recovery across analyst glove plates, ISO 5 BSC surfaces, and settling plates proves that NO viable contamination transfer pathway existed.
              4. *Closed Transport*: Materials transported in disinfected, lidded bins on carts.
              5. *Analyst Compliance*: Strict adherence to MICRO-SOP-9 and MICRO-SOP-44 with zero deviations.
    22. *Balanced Dual-Page Visual Layout & Font Harmonization Standard (多页叙述段落均衡分配与字号完全对齐规范):*
        *   **The Multi-Page AcroForm Dilemma:** In PDF AcroForms, text boxes on separate pages (Text Field50 on Page 4 and Text Field51 on Page 5) cannot dynamically flow or bridge text across pages. Cramming 11 paragraphs (~5,260 characters) into Page 5 while leaving Page 4 with only 5 paragraphs (~2,800 characters) forces Page 5 to shrink drastically to 7.325pt while Page 4 inflates to 12.0pt with an unsightly bottom void, creating an unhuman, jarring appearance.
        *   **The 50/50 Balanced Character Distribution Mandate:**
            - Narrative paragraphs MUST be divided symmetrically by character volume (~4,500 characters on Page 4 and ~4,300 characters on Page 5).
            - **Page 4 (Testing Operations & Facility Baseline):** Contains paragraphs p6..p12 (Sample receipt, direct inoculation in CR114B / BSC E001316, incubation & aliquoting in CR114A / BSC E001798, Celsis RLU readings, Subculture & Differential Staining 0 CFU, Media expiry & controls, Facility Monthly Cleaning of CR114 on 30 Aug 2026, and Introduction to Tables 2 & 3).
            - **Page 5 (Environmental Investigation & Quality Defense):** Contains paragraphs p13..p21 (Processing EM review, Aliquoting EM review, Consolidated ISO 5 EM defense, Lack of contamination pathway, Analyst interview & cleaning compliance, Sample batch & cross-contamination review, Six-month client census, Lot history, and Final human conclusion).
        *   **Harmonized Font Size (text_fontsize = 10.3pt):**
            - Both Page 4 and Page 5 MUST be set to the identical, natural font size 10.3pt.
            - Both pages fill the available field box cleanly from top to bottom with ~27-29pt of natural bottom margin, completely eliminating bottom voids, preventing text truncation, and delivering a 100% natural, human-typed appearance.

    23. *Celsis Narrative Precision & 6-Month Analyte History Standard (Celsis 叙述精简与特定分析物历史规范):*
        *   **Omission of Negative TSB Numerical Cutoff Sentence (阴性管数值细节精简):**
            - In Paragraph 8 (Celsis RLU reading summary), do NOT include detailed negative RLU and cutoff numbers for the negative TSB bottle (e.g. omit: "The corresponding TSB sample container ETX-XXXXXX-XXXX tested negative with X RLU, below the TSB cutoff of X RLU and negative control of X RLU.").
            - Simply state: "All other FTM and TSB sample bottles in the same batch tested negative."
        *   **Omission of Lot Retesting Submission Sentence (批号复检历史段落剔除):**
            - Do NOT include generic statements regarding whether the lot was submitted for retesting (e.g. omit: "A review of the lot history shows that there have been no additional submissions of sample lot X for retesting..."). Omit this paragraph entirely.
        *   **Locked Standard Phrasing for 6-Month Analyte History (6个月特定分析物历史标准句式):**
            - Standard template syntax:
              `Analyzing a 6-month sample history for [Client Name] indicates, this specific analyte "[Analyte / Sample Name]" has had no prior failures using the Celsis Sterility testing during this period.`
            - Concrete example:
              `Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "MOTs-C 10 MG/ML (5 ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.`
    24. *AcroForm PDF Zero-Overflow & Text Truncation Elimination Standard (PDF 表单全字段零截断与杜绝 '+' 按钮规范):*
        *   **Root Cause of Adobe Acrobat `+` Overflow Icon:** In PDF AcroForms, multiline text fields set to `text_fontsize = 0.0` (Auto) cause PyMuPDF's appearance stream (`/AP`) generator to default to `/Helv 12 Tf` (12pt font). In constrained table cells (e.g. Section B comments with heights of 15.8pt or 20.0pt), 12pt line leading (~13.4pt) pushes lines 2 and 3 into negative coordinates outside the widget clip box (`1 1 178.17 rh re W n`). Adobe Acrobat detects that the `/AP` stream exceeds the bounding box, truncates the visible rendering at line 1, and displays a black square with a `+` symbol at the bottom right.
        *   **Field-by-Field Calibrated Font Sizing:** All multiline fields must be assigned their exact auto-fit font sizes before calling `w.update()`, matching official signed QA baselines (e.g., `OOS-261877`):
            - Page 1 Section B Comments: `Text Field13` (4.825pt), `Text Field14` (8.5pt), `Text Field15/16` (9.0pt), `Text Field17` (6.725pt), `Text Field18` (4.825pt), `Text Field19/20` (9.0pt), `Text Field21` (5.85pt).
            - Page 1 Headers: `Text Field3` (Analyst Box: 6.225pt, all 4 analysts visible), `Text Field7` (Description of Incident: 8.0pt, 3 lines fit), `Text Field1` (Test Name: 8.0pt, 2 lines fit in 20pt), `Text Field4` (MOTs-C: 8.0pt).
            - Page 2 Equipment & Reagents: `Text Field22` (Reagent lots: 4.35pt, 14 lines fit), `Text Field23` (Reagent exps: 4.75pt, 14 lines fit), `Text Field32/33` (CR/BSC & Cal dates: 4.825pt, 2 lines fit in 15.8pt), `Text Field43/44` (Incubators & Sensors / Cal dates: 7.05pt, double spacing preserved).
            - Pages 3, 4, 5: `Text Field49`, `Text Field50`, `Text Field51` (9.2pt, natural full-page flow).
        *   **Zero Artificial Spacing / Trailing Returns:** Trailing newlines (`\r \r \r`) and trailing spaces must be stripped from field values (`.rstrip(' \r\n')`) to prevent unnecessary line increments that trigger clipping.
        *   **Mandatory Bounding Box Inspection Gate:** Automation scripts must parse generated `/AP` streams and verify that cumulative baseline vertical coordinates satisfy `min_y >= 0` across all text fields, guaranteeing 100% text visibility and complete absence of the `+` button in Adobe Acrobat.
    25. *Strict 1-to-1 Transcription Rule (纯复制粘贴转录原则：只搬运内容，不擅改字号与格式，顺其自然):*
        *   **User Direct Mandate:** 在将草稿 PDF（如 `OOS-XXXXXX ... .pdf`）转录至正式官方表单（如 `CORP-FORM-21 ... .pdf`）时，**仅执行严格的内容与勾选状态 1:1 纯复制粘贴（Direct Copy-Paste）**。
        *   **Strict Hands-off on Fonts & Formatting:** 严禁在转录脚本中主动去计算、改动、微调或覆盖字段的字号 (`text_fontsize`) 或格式。打多少字就直接赋什么值，不要管留不留白，保持表单原生默认状态，顺其自然，不进行任何额外的格式干预。
    26. *Strict Prohibition of Ampersand `&` & Informal Symbols (严禁使用 `&` 等非正式特殊符号铁律):*
        *   **Zero Ampersand Rule:** In all cGMP formal reports, investigation forms (like `CORP-FORM-21`), comments, standalone tables, and narrative texts, **NEVER use `&`** (ampersand) as a conjunction.
        *   **Full Word Spellout:** Always spell out `, and ` or ` and ` in full English sentences.
            - *Analyst Lists*: `"Yes, analysts Andrew Carrillo, Gabrielle Surber, America Alanis, and Cuong Du were comprehensively interviewed."` (NEVER `Alanis & Cuong Du`).
            - *Section Headers & Reagents*: `"Celsis Reagents and Kits:"` (NEVER `Reagents & Kits`).
        *   **Zero Ampersand in Bacterial Tables:** In Table 2 / Table 3, DO NOT concatenate multiple organisms with `&`. Separate them with a blank line.
    27. *Hyperlink Line Spacing & Vertical Baseline Alignment Invariance (超链接段落行距与水平基线绝对对齐铁律):*
        *   **Absolute Ban on `line="360"`:** When injecting OpenXML `<w:hyperlink>` into table cells in Word/Docx, **NEVER inject `<w:spacing w:line="360" w:lineRule="auto"/>`**! Extra line spacing expands the single-line bounding box and causes Word's vertical alignment (`w:vAlign="center"`) to center an inflated box, misaligning the hyperlinked Sample ID relative to adjacent cells (`1 CFU`, `None`).
        *   **Preserve Native Cell Paragraph Properties:** Keep standard paragraph spacing (no extra line spacing), inherit cell paragraph properties `<w:jc w:val="center"/>`, exact 7pt font (`w:sz="14"`), blue color (`#0000FF`), and single underline (`<w:u w:val="single"/>`).
        *   **Sub-Pixel Vertical Alignment:** In exported PDFs, the Sample ID hyperlinked text must match the exact vertical baseline (`y0`, `y1`) of adjacent cells down to 0.001pt precision.
    28. *Microbiological Binomial Nomenclature Formatting Standard (微生物拉丁双名法全斜体与隔行规范):*
        *   **Mandatory Italics for Species (*Genus species*):** All bacterial genus and species names (*Corynebacterium ureicelerivorans*, *Mycobacterium grossiae*, *Micrococcus luteus*, *Staphylococcus epidermidis*, etc.) **MUST ALWAYS BE ITALICIZED** (`<w:i/>`, `<w:iCs/>` in Word; `<i>...</i>` or markdown `*...*` in docs).
        *   **Multi-Organism Blank Line Separation:** When multiple organisms are recovered from the same sampling site/plate, display each organism on its own line(s), separated by an empty blank line in between. Do NOT concatenate with `&` or commas.
    29. *Sterility Sample Inoculation Volume Phrasing Standard (per media type 培养基接种体积术语铁律):*
        *   **Mandatory Wording (`per media type`):** In direct inoculation sterility testing (Celsis, USP <71>, Scan RDI), when the testing protocol divides the total sample volume across multiple containers/jars per media type (e.g., five 300 mL FTM jars and five 300 mL TSB jars), the narrative MUST strictly state:
            `"Direct inoculation was performed by adding [X] mL of sample per media type, using [N] [size] [Medium 1] jars and [N] [size] [Medium 2] jars."`
            - Example: `"Direct inoculation was performed by adding 25 mL of sample per media type, using five 300 mL FTM jars and five 300 mL TSB jars."`
        *   **Strict Prohibition of `per media container`:** **NEVER** write `"per media container"`. Stating "per media container" erroneously implies that each individual jar received 25 mL (which would multiply the total sample volume tested to 250 mL, creating a serious factual contradiction against laboratory batch records and client submissions).
    30. *Suitability Method Record Notation & Positive Media Reading Conciseness Standard:*
        *   **Suitability Record Notation:** When citing suitability method records in investigation narratives, always format with a colon after `record`: `"suitability method record: [RECORD_NUM]"` (e.g., `"suitability method record: 2601120497"`).
        *   **Positive Container Redundancy Elimination:** In the analytical reading summary sentence, do NOT append sub-container tracking IDs (e.g. `(ETX-XXXXXX-XXXX-X/N)`) after naming the positive media jar (state `"sample ETX-XXXXXX-XXXX was found to yield a positive reading in one of the 300 mL FTM media jars."`). Redundant sub-container IDs clutter narrative flow and are already explicitly captured in Table 1.
    31. *Sterility Positive Media Jar Count & Sequence Verification Gate (无菌检测阳性培养基瓶数与瓶序多重绝对校验铁律):*
        *   **Raw Container Count Primacy (原始瓶数绝对守恒原则):** When raw laboratory notification emails or batch sheets state `$N \times \text{[size]} \text{ [media]}$` (e.g., `2 x 300mL FTM`), the integer $N$ is the immutable ground truth. Table 1 Column `Media with microbial growth` MUST strictly match the exact count: `$N \times \text{[size]} \text{ [media]}$` (e.g., `2 x 300mL FTM`). NEVER reduce, abbreviate, or round down to `1 x`.
        *   **Absolute Ban on Fraction Misinterpretation (严禁将多瓶序号压缩为分数导致单瓶误判):** Phrases indicating specific jar positions (e.g. `(3rd and 5th jars)` or `jars #3 and #5`) MUST NEVER be shorthand-compressed into notations like `3/5`. The AI and scripts MUST NEVER interpret `3/5` as "bottle 3 of 5". Multi-jar enumerations must be explicitly preserved as distinct positive containers.
        *   **Grammatical & Contextual Concordance Across All Narrative Sections (全篇前后文单复数严格一致性准则):** When $N \ge 2$, Narrative Paragraph 8 MUST state `"sample [ETX] was found to yield positive readings in [word(N)] of the [size] [media] media jars ([explicit jars])"`, and all references to containers MUST use plural nouns (`positive FTM sample bottles`, `duplicate reading tubes`, `bottles were submitted`).
        *   **Pre-Delivery Raw Data Reconciliation Gate (交付前原始邮件与成品强制逐字穿透对账闸门):** Before declaring completion of ANY OOS report, the AI MUST execute an automated cross-reconciliation check between user raw prompt text and the generated tables/narratives. If any container count mismatch is detected, execution MUST immediately abort and trigger self-correction.
    32. *Analyst Name Accuracy & ScanRDI Prepping Analyst Mapping (分析员真实法定全名与 ScanRDI 预处理分析员标准):*
        *   **Legal / Full Name Precision (Mukyung Jang vs. Min Jang):** For initials MJ in Scan RDI sample preparation, the full legal system name is **Mukyung Jang (MJ)** (NEVER shorthand or truncated Min Jang).
        *   **Global Concordance:** When Mukyung Jang is updated, the name MUST be synchronized across Section B comments (Text Field13), Page 3 narrative (Text Field49), and master Word report blocks.
        *   **Two-Page Narrative Flow & Page 5 N/A Invariance:** When narrative content fits symmetrically across Page 3 (Text Field49, fs=8.75pt, ~3,540 chars) and Page 4 (Text Field50, fs=9.2pt, ~3,375 chars), Page 5 Text Field51 is marked as 'N/A QYC [Date]' to maintain clean visual balance. Page 6 Text Field54 (Lab Manager) remains empty ('') for supervisor review.
    33. *Scan RDI Changeover (S/O) Bench Operator Precision & Clean EM Attribution Standard:*
        *   **S/O Operator Primacy:** In Scan RDI OOS investigations, the Changeover Analyst is the actual operator who performed the changeover procedure in the biological safety cabinet as signed on the physical bench logbooks. For OOS-261967, bench records confirm Changeover was performed by **Karla Silva (KSM)** in BSC E001937, while Testing was performed by **Varsha Subramanian (VV)** in BSC E001319.
        *   **Multi-Line Role Layout in Text Field 3:** Page 1 Text Field3 must format all 4 roles distinctly:
            ```text
            Prepping Analyst: \rMukyung Jang (MJ)\r \rProcessing Analyst: \rVarsha Subramanian (VV)\r \rChangeover Analyst: \rKarla Silva (KSM)\r \rReading Analyst: \rSonal Uprety (SU)
            ```
        *   **Clean EM Attribution:** Both Testing (VV in BSC E001319) and Changeover (KSM in BSC E001937) personnel touch plates, surface contact plates, and settling plates demonstrated 100% absence of microbial recovery (No Growth / 0 CFU). Table 2 explicitly captures separate rows for Testing and Changeover (S/O) across Personnel, Surface, and Settling sites, accurately reflecting their respective operators (VV and KSM) and BSC equipment IDs (BSC E001319 and BSC E001937). Narrative texts reflect clean critical zones with no positive recovery during testing or changeover.
    34. *Celsis Sterility Reviewer Standards & QA Golden Feedback (Robin Sharma / RS Review Feedback for OOS-262080):*
        *   **Method Suitability Matching Negative Control Rule (适用性添加物负对照强制闭环准则):** When suitability testing specifies adding an additive or neutralizer (e.g., 1g BSA in 100 mL PBS, adding 5 mL to each jar), the narrative MUST explicitly state: `"Accordingly, a negative control with the same modifications was made."` This proves that the modified media system and additive remained sterile and did not introduce the contamination.
        *   **Individual RLU Reporting for Multiple Positive Bottles (多瓶阳性独立 RLU 数值报告铁律):** When multiple containers/bottles test positive, do NOT report only a single combined average RLU across all bottles. MUST report each positive bottle individually with its respective position and reading: `"The confirmed average Relative Luminescence Units (RLU) from the duplicate reading tubes, originating from the positive FTM sample bottles, yielded 4,202 RLU (Bottle 3) and 7,190 RLU (Bottle 5), both exceeding the positive cutoff of 2,955 RLU, where the FTM negative control was 985 RLU."`
        *   **Colony Count Reconciliation in EM Active Air (活菌计数与菌种数量绝对统一准则):** When an active air sample recovers multiple bacterial isolates (e.g., *Corynebacterium ureicelerivorans* and *Mycobacterium grossiae* on `ETX-260914-0487`), the total colony count must accurately reflect all isolates (`2 CFU (ISO 8 114)`, NEVER understated as 1 CFU).
        *   **Standard Subculture Negative Terminology ("Microbial growth could not be recovered"):** In Table 1 Column `Microbial ID`, when subculture yields no growth for a positive sterility reading: MUST write `"Microbial growth could not be recovered"` (NEVER colloquial `"No growth was obtained"`).
        *   **Universal EM Weekly Column Header ("Week of Testing"):** In Tables 2 and 3, Column `Day /Week(s)` for weekly environmental monitoring rows MUST STRICTLY be `"Week of Testing"` (NEVER `"Week on Testing Date"` or `"Week on Testing Date."`). Root templates `tables for celsis.docx` and `tables for 71.docx` have been updated to enforce this permanently.
    35. *Celsis OOS-262017 Table Finalization & Aliquoting Weekly EM Integration (Revive Rx Pharmacy, E00927):*
        *   **Complete Bracketing Resolution:** Resolved all daily bracketing rows in Table 2 (Processing Phase, 24Aug26 by ES) and Table 3 (Aliquoting Phase, 31Aug26 by ALA) to `"No growth"` / `"N/A"` / `"None"` based on comprehensive EM census review.
        *   **Weekly EM Active Air Integration (04Sep26 under ETX-260914-0487):** In Table 3 Row 13 (Aliquoting Weekly Active Air), incorporated the ground truth recovery of `2 CFU (ISO 8 114)` sampled on `04Sep26` by analyst `ISS` under plate `ETX-260914-0487` (with active clickable hyperlink). Recovered isolates are formatted as *Corynebacterium ureicelerivorans* and *Mycobacterium grossiae* (both italicized, separated by empty blank line, 0 ampersands).
        *   **Visual Ergonomics & Zero Ampersands:** Enforced 100% cell centering, consistent 7pt Times New Roman, zero `&` symbols, active clickable hyperlinks across Table 1 and Table 3, clean page break before Table 3, and flawless Word COM PDF export.
    36. *Anti-Robotic Human Narrative Standard & Zero Special Symbol Mandate (极致拟人化自然写作与全面清除机械符号永久铁律 - 🚨 终极合规规范):*
        *   **🚨 机械化写作零容忍准则 (Zero Tolerance for Robotic Syntax):** 严禁任何具有机械 AI 感的符号、括号式数据堆砌、以及段落中的冒号小标题。调查报告必须呈现 100% 资深人类质量科学家（Senior QA Scientist / Microbiologist）的流利英语行文质感。
        *   **1. Zero Ampersand (`&`) Everywhere (绝对零 `&` 铁律):**
            - 在所有动态变量、分析员访谈文本、表单注释、耗材/设备栏目及正文段落中，**严禁使用 `&` 作为连词**。
            - 必须全部拼写为 `, and ` 或 ` and `。
            - 例：`America Alanis, and Cuong Du`（严禁 `Alanis & Cuong Du`）；`Celsis Reagents and Kits`（严禁 `Reagents & Kits`）。
        *   **2. Complete Elimination of Bracketed Data Dumps (严禁括号式数据堆砌，全面转化为流利完整句子):**
            - 人类调查员在书写正式报告时，从不把数据装在密集的括号 `()` 里面。所有技术参数、读数、批号、日期、样本号必须自然融入完整的英语主谓宾句子中。
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
        *   **3. Zero Robotic Colons in Continuous Prose (段落中彻底剔除伪小标题冒号):**
            - 严禁在连贯叙述中插入像 `Evaluation of Environmental Monitoring Results:` 这样的突兀冒号。
            - 必须用自然的引导从句过渡：`Regarding the evaluation of environmental monitoring results, it is important to note that...`。
            - 引用适用性记录编号时，严禁加冒号：`suitability method record ETX-251218-0432`（严禁 `suitability method record: ...`）。
    37. *Scientific Units Preservation & Dual Date Format Standard (DDMMMYY 与 MMM YYYY 规范及科学专业单位守恒铁律 - 🚨 终极标准):*
        *   **1. Professional Scientific Units Preservation (专业科学符号与单位绝对保留，严禁误伤硬改成英文单词):**
            - 专业符号与度量衡单位（如 `%`, `%CV`, `CV%`, `℃`, `CFU`, `CFUs`, `RLU`）属于 cGMP 实验室的标准学术表达，**绝对不能当作“特殊符号”被误伤硬改成英文单词**！
            - `%` / `%CV` / `CV%`：必须保留符号（如 `2%`, `%CV < 30%`, `CV% of 3%`，严禁改成 `percent` 或 `percent CV`）。
            - `℃`：保留摄氏度单位符号（如 `30 to 35 ℃`, `20 to 25 ℃`，严禁写成 `degrees Celsius` 或 `degrees C`）。
            - `CFU` / `CFUs`：保留微生物菌落计数单位。
            - `RLU`：保留荧光读数单位。
            - “少用特殊符号”仅针对非正式文本符号（如 `&` 代替 `and`、段落中滥用伪小标题冒号 `:`、或把主谓宾数据丢进密集括号 `(...)`）。
        *   **2. Dual Date Format Standard in Narrative Prose (正文叙述日期双轨制铁律):**
            - **具体日期（有 DD）**：**必须使用 `DDMMMYY` 紧凑格式**（如 `21Aug26`, `24Aug26`, `31Aug26`, `12Jul26`, `16Aug26`, `25Aug26`, `28Aug26`, `01Sep26`, `04Sep26`）。中间严禁带空格，严禁写成 `24 Aug 2026`。
            - **宽泛日期（无 DD，仅年月）**：**必须使用 `MMM YYYY` 格式**（如 `Jul 2026`, `Aug 2026`, `Sep 2026`, `Jan 2027`）。
        *   **3. Pruning & High-Level Narrative Standard (去粗取精与正文高级叙述原则):**
            - 耗材批号/效期（如 Tween 80 批号、TSB/FTM 批号）已在 Page 2 耗材栏详细列出，正文段落无需重复罗列其批号与效期。
            - 仪器设备（如 Celsis 读数仪）正文中直接写作 `Celsis E002222`，无需长串修饰 `Advance 2, equipment ID E002222`。
            - 历史记录（6个月记录）直奔主题，直接指明该特定分析物（Tesamorelin 12 mg per Vial）的历史阳性次数与 OOS 编号（OOS-261500 和 OOS-261878），剔除无意义的批号复测套话。
            - 同批交叉污染段落精炼为明确的 3 句话，无需冗余罗列 6 个同批阴性 ETX 编号。
        *   **4. Logical 3-Page Flow (三页黄金结构):**
            - Page 3 (`Text Field49`): 实验前准备、洁净室、接种与分装（P1..P7）
            - Page 4 (`Text Field50`): Celsis 读数结果、菌种确认、效期、月度清洁与表格引言（P8..P12）
            - Page 5 (`Text Field51`): 环境监测调查、交叉污染排除、6个月历史审查与最终结论（P13..P21）
    38. *USP <71> OOS-262098 Master Production Package (Solyn LLC, E75000):*
        *   **Full Production Pipeline:** Finalized and delivered full production package for sample `ETX-260902-0505` (`GLP3R/Cagrilinitide`, Lot: `2608-216`) including Standalone Tables DOCX/PDF, Master Word Report (`OOS-262098 Solyn LLC (E75000) - USP71.docx`), and Official 7-Page PDF Form (`OOS-262098 Solyn LLC (E75000) - USP71.pdf`).
        *   **Lab Scheduling & Holiday Logic:** Corrected Friday (`04Sep26`) testing follow-up date to skip weekend (`05Sep26` Sat & `06Sep26` Sun) and Labor Day (`07Sep26` Mon) to Tuesday `08Sep26` (ES) for all Table 2 bracketing rows.
        *   **Weekly Active Air & Ground Truth Hyperlink:** Integrated PRD ground truth for 04Sep26 active air under `ETX-260914-0487` (analyst `ISS`, `2 CFU (ISO 8 114)`, clickable URL to `fd3G2StZClcy1TP2ES6BLw__`, *Corynebacterium ureicelerivorans* & *Mycobacterium grossiae*).
        *   **Confirmed Microbial Identification:** Reconciled preliminary broth Gram (-) reading with definitive molecular sequencing under `ETX-260910-0290` as *Microbacterium sp. PM5* (`Gram (+) rods`) from pure TSA & SDA subcultures.
        *   **Official PDF Transcription:** Transcribed onto `CORP-FORM-21 - P1 04 SEP 2026.pdf` enforcing Rule 9 Incubator spacing (`\r \r` with sensor on same line), Rule 11 calibrated auto-fit font sizes, and appended Table 1 & 2 as Page 7.
    39. *USP <71> OOS-262098 Phase II Full Investigation Package (Solyn LLC, E75000):*
        *   **Product Contamination Root Cause Conclusive Confirmation:** Initial test under `ETX-260902-0505` failed on Day 6 in TSB (*Microbacterium sp. PM5*, `ETX-260910-0290`, Gram (+) rods, Plate A TSA & Plate B SDA). Retest under `ETX-260914-0470` (prepped by Andrew Carrillo, processed by Abayomi Odugbesi) replicated the exact failure kinetics, turning turbid positive on Day 6 in TSB, with sequencing under `ETX-260921-0498` confirming the identical strain: ***Microbacterium sp. PM5*** (`Gram (+) short rods`, Plate C TSA & Plate D SDA). Definitively refuted preliminary analyst handling error hypothesis during reconstitution and confirmed inherent batch contamination (Lot `2608-216`).
        *   **Prior Passing ScanRDI Justification (Non-Uniform Microbial Distribution):** Rigorously justified why prior rapid ScanRDI testing (`ETX-260807-0602`) passed: compounded parenteral products with low bioburden exhibit heterogeneous, non-uniform spatial distribution across vials; sampled ScanRDI vials contained zero viable cells, while USP <71> vials contained low-level cells that were enriched over 14 days in TSB, turning turbid on Day 6 twice in a row.
        *   **Phase II Universal Pipeline:** Generated full Phase II deliverables including Standalone 1-Page Tables (`Tables OOS-262098 Solyn LLC (E75000) - Phase II.docx`/`.pdf`), Master Word Report (`OOS-262098 Solyn LLC (E75000) - Phase II.docx`), and Official 6-Page PDF Form (`OOS-262098 Solyn LLC (E75000) - Phase II.pdf` & `... - QYC.pdf`).
        *   **Phase 1 & Phase 2 Form Alignment:** Updated Phase 1 Form 3.100.019.F01 Page 6 to check `Check Box87` (Yes, lab error investigated), `Check Box89` (Yes, initiate Form 3.100.019.F02), and `Check Box91` (Yes, cannot close, initiate F02). In Phase II Form 3.100.019.F02, checked `Check Box51` (Yes, root cause identified), `Check Box61` (External Phenomena - Product Bioburden), and confirmed original result valid (failing) and retest failing.
    40. *Universal Mandatory Blank Line Separation for Multi-Item Table Cells (表格多菌种/多记录单元格绝对强制空行隔开永久铁律 - 🚨 终极防线):*
        *   **🚨 终极合规背景与血的教训 (Zero Tolerance for Adjacent Unseparated Items):** 在所有 OOS 调查报告的表格中（包括 Table 1、Table 2、Table 3 以及独立表格附件），当单个单元格内存在多个条目（如多个不同菌种、多行活菌检出 Observation、多行平板 ETX 编号等）时，**每一个条目之间必须用严格且清晰的空行（Empty Blank Line / `\n\n` / 独立间距段落 `<w:p>`）隔开**！严禁将两个条目紧贴在相邻行上，任何相邻条目之间未留空行的输出均属严重交付事故！
        *   **1. 菌种鉴定栏目空行铁律 (Distinct Microbial Strains Separation):**
            - 在 `Microbial ID` / `Related Microbial ID` 单元格中，如果检出 2 个或 2 个以上的微生物菌株，**每一个菌株必须独立成行，且菌株与菌株之间必须严格插入一个独立空行**！
            - ❌ **严重违规 (相邻紧贴)**：
              *Corynebacterium sp*
              *Micrococcus luteus*
            - ✔️ **合规标准 (绝对空行)**：
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
        *   **2. 多点检出环境监测栏目横向对齐空行铁律 (Multi-Hit EM Observations & ETX Numbers):**
            - 当同一洁净区或同一监测行存在多处阳性检出时（例如 L-Suite 周检回溯出现 3 处检出）：
              - `Observation` 列中，每一个地点的检出数据必须以空行相隔：
                `6 CFUs (ISO 8 143 Sec I)`
                [空行]
                `8 CFUs (ISO 8 143 Sec II)`
                [空行]
                `2 CFUs (ISO 8 142)`
              - `EM Plate ETX Number` 列中，对应的每一个平板编号必须同样以空行相隔，并与 Observation 保持绝对严格的横向基线 1:1 对齐：
                `ETX-260929-0335`
                [空行]
                `ETX-260929-0341`
                [空行]
                `ETX-260929-0344`
        *   **3. 底层代码实现与防塌缩技术标准 (OpenXML & String Formatting Standard):**
            - 在纯文本/数据字典拼接时，多项条目一律使用 `\n\n`（双换行）连接，严禁使用单换行 `\n`。
            - 在 `python-docx` / OpenXML 底层构建单元格时，不能仅依赖可能被 Word 渲染引擎折叠的微小行高（如 `line="80"`），必须插入带有合适行距的独立空白段落 `<w:p>`，确保在 Word 和 Adobe Acrobat 中视觉上清晰可见整行空白高度。
        *   **4. 交付前自动审计闸门 (Pre-Delivery Table Line-Spacing Audit Gate):**
            - 任何脚本在交付表格前，必须执行单元格行距自检：扫描所有文本包含多行的单元格，确认非空行之间必须存在空白行间隔，杜绝任何人眼找茬被抓现行。

## Pending/Future Work
*   **Roll out Smart Justification to USP <71>:** The engine is live for Celsis and Scan RDI, but `USP71.py` still needs its underlying logic updated to utilize the 4-Step Shielding Mechanism and the new "RS Reviewed" narrative format (adjusting for its specific workflow).
*   **Template Updates:** The underlying `.docx` and `.pdf` templates need manual layout updates (by the user in Word/Acrobat) to accommodate the significantly longer narrative text before they can be perfectly auto-filled without cutoff/font-shrinking.
*   **OOS Table Generator (`generate_oos_table.py`) — Phase 2:** Core script is built and verified for all 4 OOS types (EM, Scan, Celsis, USP71). Phase 2 will standardize the AI extraction prompt/checklist for reading scanned PDF logbook images and outputting structured data dicts. The workflow is: user provides scanned PDF copy → AI reads images and extracts ETX IDs, CFU counts, dates, analysts → AI builds data dict → script fills existing `tables for XX.docx` templates → outputs final Word table. Microbial IDs provided separately by user.
