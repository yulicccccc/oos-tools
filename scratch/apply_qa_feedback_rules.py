import sys

# 1. Update .agents/AGENTS.md
with open('.agents/AGENTS.md', 'r', encoding='utf-8') as f:
    content = f.read()

# Update Dec 2026 for BSC E001316 to Jun 2027
old_bsc = 'Always explicitly list the primary processing BSC (e.g. `ISO 5 BSC E001316`) with its calibration due date (`Dec 2026`).'
new_bsc = 'Always explicitly list the primary processing BSC (e.g. `ISO 5 BSC E001316`) with its calibration due date (`Jun 2027`).'
assert old_bsc in content, "old_bsc not found in AGENTS.md"
content = content.replace(old_bsc, new_bsc)

# Add RULE 0U after RULE 0T
rule_0u = """
## 🚨🚨 RULE 0U (TOP PRIORITY): QA Auditor (Tahsin Kevenk) Cross-Document Harmonization & Powder/Peptide Reconstitution Accuracy Mandate (质保部 QA 审查跨报告数据一致性、设备日期绝对对齐与冻干多肽复溶量零臆造铁律)

**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), investigations submitted for QA Department approval (Audited by Tahsin Kevenk et al.) MUST strictly satisfy these two non-negotiable QA audit checkpoints:

1. **冻干粉与重构复溶肽类测试量/单位零脑补准确性 (Reconstituted Peptide Sample Amount & Unit Accuracy - QA Defect Zero Tolerance)**:
   - 凡送检样品为冻干粉、小容量西林瓶注射剂或需要复溶的肽类药物（如 GLP3R/Cagrilinitide, Tirzepatide, Sermorelin, Retatrutide 等），**严禁在叙述中擅自写出虚假质量重量单位**（🚨 **绝对严禁写出 `"60 gm filtered across 4 vials"` 或 `"45 gm filtered"` 等凭空捏造的质量单位**！）
   - 必须严格核对 EagleTrax 原始工单（Raw Data）与样品描述，统一按实际复溶规程叙述：
     `"via membrane filtration (2mL Bacteriostatic water per vial each containing [Analyte] in [Concentration]). Half the contents of each reconstituted vial was then dispensed directly from the container into sterile canisters placed on a vacuum manifold for membrane filtration."`
   - **跨报告吻合铁律**: Phase I 与 Phase II 在样品处理数量、浓度、复溶溶媒（Bacteriostatic Water）及分液过滤比例上**必须完全统一、严丝合缝**！

2. **Phase I 与 Phase II 设备校准与洁净室认证到期日绝对对齐标准 (Equipment Calibration & Cleanroom Certification Due Date Harmonization)**:
   - Phase I 与 Phase II 涉及的相同设备与洁净室，其 Calibration / Certification Due Date **必须 100% 绝对一致，严禁出现矛盾**：
     - **`ISO 5 BSC E001316`**：Calibration Due Date 必须统一为 **`Jun 2027`**（🚨 严禁在 Phase I 写 Dec 2026 而在 Phase II 写 Jun 2027！两者必须完全统一为 `Jun 2027`）！
     - **`Cleanroom CR114 (E001736)`**：Certification Due Date 必须统一为 **`Dec 2026`**（🚨 严禁在 Phase I 写 Jan 2027 而在 Phase II 写 Dec 2026！两者必须完全统一为 `Dec 2026`）！
     - **`Cleanroom CR115 (E001737)`**：Certification Due Date 必须统一为 **`Dec 2026`**（严禁写 June 2027）！
     - **`ISO 5 BSC E001314`**：Calibration Due Date 必须统一为 **`Jun 2027`**（与同套间设备对齐）！

3. **QA 审核员花名册与三级审查认知 (QA Auditor Identity & Three-Tier Signoff Protocol)**:
   - **Tahsin Kevenk** 为质保部（Quality Assurance / QA）核心审核人，重点审查跨报告数据一致性、单位准确性及设备校准到期日！
   - 牢记三级流转：Prepared by: **Qiyue Chen (QYC)** ✍️ $\to$ Reviewed by: **Robin Seymour (RS)** (Micro Lab Supervisor) ✍️ $\to$ Approved by (QA): **Tahsin Kevenk / QA Department** ✍️。

---
"""

# Find insertion point: after RULE 0T
target_anchor = "## 🚨🚨 RULE 0T (TOP PRIORITY): Sterile Suite & Biosafety Cabinet Analyst Assignment Ground Truth (无菌套间、洁净室与分析员常态化主工位权威映射)"
assert target_anchor in content, "target_anchor not found in AGENTS.md"
# We can insert RULE 0U right before RULE 0T
content = content.replace(target_anchor, rule_0u + "\n" + target_anchor)

with open('.agents/AGENTS.md', 'w', encoding='utf-8') as f:
    f.write(content)
print("Updated .agents/AGENTS.md successfully!")

# 2. Update PRD.md
with open('PRD.md', 'r', encoding='utf-8') as f:
    prd_content = f.read()

# Update Dec 2026 for BSC E001316 in Section 63 Item 1
old_prd_bsc = "除了培养箱与传感器外，必须明确列出主处理生物安全柜 `ISO 5 BSC E001316` 及其校准到期日 `Dec 2026`。"
new_prd_bsc = "除了培养箱与传感器外，必须明确列出主处理生物安全柜 `ISO 5 BSC E001316` 及其校准到期日 `Jun 2027`（与 Phase II 绝对对齐，经 QA Tahsin Kevenk 审核定稿）。"
if old_prd_bsc in prd_content:
    prd_content = prd_content.replace(old_prd_bsc, new_prd_bsc)

# Add Section 65 to PRD.md
sec_65 = """
65. **QA Official Review & Cross-Document Harmonization Standards (Tahsin Kevenk QA Review 质保核准定稿标杆与两大黄金对齐铁律 - 🚨 终极锁定):**
    *   **标杆案卷与审查来源**: 基于质保部（Quality Assurance / QA）审查员 **Tahsin Kevenk** 对 `OOS-262098 Solyn LLC (E75000) - Phase I & Phase II` 提出的法定审核意见，以及最终由 QYC 调查起草、RS 审阅、QA 核准的最新终极已签版本（`... - RS Reviewed - Signed by QYC-RS - Updated.pdf`），确立以下两大跨案卷数据对齐铁律：
    *   **1. 冻干粉/多肽复溶量与单位零脑补真实性铁律 (Reconstituted Peptide Sample Amount & Unit Accuracy):**
        - 在 Phase II Page 4 及正文叙述中，凡样品为冻干粉小西林瓶（如 GLP3R/Cagrilinitide 25mg/5mg），**严禁擅自使用虚假质量单位**（严禁写 `"60 gm filtered across 4 vials"` 或 `"45 gm filtered"`）！
        - 必须与 Phase I 保持 100% 严密一致，如实按照复溶规程叙述：
          `"...via membrane filtration (2mL Bacteriostatic water per vial each containing GLP3R/Cagrilinitide in 25mg/5mg concentration). Half the contents of each reconstituted vial was then dispensed directly from the container into sterile canisters placed on a vacuum manifold for membrane filtration."`
    *   **2. Phase I 与 Phase II 设备校准与洁净室认证到期日绝对对齐 (Cross-Phase Equipment & Cleanroom Due Date Harmonization):**
        - 两份报告中所列之同一设备与洁净室，其校准/认证到期日必须严密吻合，绝不允许出现冲突：
          - **`ISO 5 BSC E001316`**: Calibration Due Date 统一为 **`Jun 2027`**（Phase I 曾误填为 Dec 2026，现终极更正并锁定为 `Jun 2027`）。
          - **`Cleanroom CR114 (E001736)`**: Certification Due Date 统一为 **`Dec 2026`**（Phase I 曾误填为 Jan 2027，现终极更正并锁定为 `Dec 2026`）。
          - **`Cleanroom CR115 (E001737)`**: Certification Due Date 统一为 **`Dec 2026`**。
          - **`ISO 5 BSC E001314`**: Calibration Due Date 统一为 **`Jun 2027`**。
    *   **3. QA 审核员架构认知 (QA Reviewer Role in 3-Tier Hierarchy):**
        - 质保审查人（如 **Tahsin Kevenk**）独立于实验室主管（Robin Seymour），对流转至第三级的报告进行最终合规性审查。起草阶段必须自主完成上述双重交叉比对，做到零缺陷提交。
"""

prd_content += "\n" + sec_65

with open('PRD.md', 'w', encoding='utf-8') as f:
    f.write(prd_content)
print("Updated PRD.md successfully!")

# 3. Update PROJECT_STATE.md
with open('PROJECT_STATE.md', 'r', encoding='utf-8') as f:
    state_content = f.read()

state_update = """
### Recent QA Review & Harmonization Updates (Tahsin Kevenk QA Review on OOS-262098):
- **Feedback Source**: Tahsin Kevenk (QA) reviewed OOS-262098 (Phase I & Phase II) and flagged two cross-document discrepancies:
  1. Phase II Page 4 Sample Amount: Erroneously stated "60 gm filtered across 4 vials" instead of the exact 2 mL Bacteriostatic water reconstitution narrative.
  2. Equipment Calibration / Certification Due Dates: Phase I had Dec 2026 for BSC E001316 and Jan 2027 for CR114, whereas Phase II had Jun 2027 and Dec 2026.
- **Harmonized & Approved Ground Truth**:
  - Sample Amount Narrative: Replaced with exact reconstitution narrative matching Phase I (`2mL Bacteriostatic water per vial each containing GLP3R/Cagrilinitide in 25mg/5mg concentration`).
  - Equipment Dates Harmonized:
    - `ISO 5 BSC E001316`: Calibration Due Date = **`Jun 2027`**
    - `Cleanroom CR114 (E001736)`: Certification Due Date = **`Dec 2026`**
    - `Cleanroom CR115 (E001737)`: Certification Due Date = **`Dec 2026`**
    - `ISO 5 BSC E001314`: Calibration Due Date = **`Jun 2027`**
  - Updated signed deliverables verified on Desktop:
    - `OOS-262098 Solyn LLC (E75000) - USP71 -RS Reviewed - Signed by QYC-RS - Updated.pdf`
    - `OOS-262098 Solyn LLC (E75000) -  USP 71 - RS Reviewed - Phase II - Signed by QYC-RS - Updated.pdf`
- **Locked Rules**: Implemented `RULE 0U` in `.agents/AGENTS.md` and Section 65 in `PRD.md`.
"""

state_content += "\n" + state_update

with open('PROJECT_STATE.md', 'w', encoding='utf-8') as f:
    f.write(state_content)
print("Updated PROJECT_STATE.md successfully!")
