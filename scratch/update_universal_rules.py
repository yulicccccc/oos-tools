# -*- coding: utf-8 -*-
import sys

rule_0r_agents = '''## 🚨🚨 RULE 0R (TOP PRIORITY): Universal 3-Tier OOS Review Hierarchy & Supervisor (RS) vs QA Role Distinction (全 OOS 模块三级签名流转与主管 RS vs 质保 QA 角色法定铁律)

**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), strictly maintain the official 3-tier organizational hierarchy and review roles:
1. **Tier 1 — Prepared by: (Initiating Analyst / Report Author)**:
   - **Person**: **Qiyue Chen (QYC)** (Microbiologist / 调查报告起草人).
   - **Action**: Gathers raw data, writes report narrative, fills AcroForms, signs `Signature_0`.
2. **Tier 2 — Reviewed by: (Microbiology Lab Supervisor)**:
   - **Person**: **Robin Seymour (RS)** (Microbiology Supervisor / 直接实验主管).
   - **Action**: Direct lab supervisor review. RS checks technical lab facts, edits narrative, disposes Section E (`N/A RS [Date]`), and signs `Signature_1` AFTER the author signs.
3. **Tier 3 — Approved by (QA): (Independent Quality Assurance Department)**:
   - **Department**: **Quality Assurance (QA)** (质保部独立终审).
   - **Action**: QA conducts final compliance review and closure sign-off (`Signature_2`) AFTER lab internal signatures. **QA does not inspect the report while under Supervisor review!**
4. **Mandatory Routing Order**:
   `Qiyue Chen (Prepared by) ✍️` -> `Robin Seymour (Reviewed by) ✍️` -> `QA Department (Approved by) ✍️`.

## 🚨🚨 RULE 0S (TOP PRIORITY): Universal Sterility Investigation Standards from Robin Seymour (全 OOS 模块复溶详述、设备栏 "Phase I Summary"、CR115 校准期、冲突抗辩两段式、根因与结语极简法则)

**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`):
1. **Mandatory Reconstitution Process Narrative (冻干粉/多肽复溶全过程必写)**:
   - Whenever testing lyophilized vials, powders, or peptide products across ANY test module, narrative MUST detail the reconstitution process (diluent name, volume added, analyte concentration, and aliquot portion filtered), NEVER vaguely write a filtration weight or bottle count alone!
2. **Equipment Section Reference & Primary BSC List (设备栏 "Phase I Summary" 与主生物安全柜必录)**:
   - In Page 2 equipment tables across all forms, environmental control references must be written as `"Phase I Summary"` (NEVER `"See Phase I Summary"`).
   - Always explicitly list the primary processing BSC (e.g. `ISO 5 BSC E001316`) with its calibration due date (`Dec 2026`).
3. **Cleanroom Suite 115 Certification Due Date (CR115 校准期统一)**:
   - CR 115 (`E001737`) calibration due date is strictly **`Dec 2026`** (identical to CR 114 / `Dec 2026`), NEVER June 2027.
4. **Robin's Golden 2-Paragraph Defense for Prior Passing Rapid Test vs Current Failure (先前快速法合格与当前检验不合格冲突时的两段式抗辩金标准)**:
   - Paragraph 1: State that review determined *no approved method suitability test* was available to support the validity for microbial growth recovery using the rapid method for this specific sample.
   - Paragraph 2: State that discrepancy is attributable to *non-homogeneous distribution of microbial contamination at very low levels* within the compounded batch, and conclude that *given the absence of a qualified method suitability study, the prior passing result cannot be considered directly comparable to the current test outcome*.
5. **Phase II Section D Root Cause Simplification (根因分析极简核心)**:
   - In Section D root cause statement, DO NOT write long speculative essays. Strictly focus on confirmed recovery of identical isolate during retest and state the original failure is confirmed valid:
     `"Inherent sample contamination within Lot [Lot #], supported by recovery of the same organism ([Microbial ID]) during retesting. The original [Test Method] sterility failure was confirmed as valid."`
6. **Investigation Conclusion (调查结语双有效性判定)**:
   - Explicitly state inherent product contamination and that *both the initial failing result and the retest failing result are considered valid*:
     `"Therefore, inherent product contamination in Lot [Lot #] is confirmed, both the initial failing result for [Initial ETX] and the retest [Retest ETX] is considered valid."`
7. **Section E Handling (Section E 权限绝对留空)**:
   - Investigators must leave Section E completely empty during drafting. Direct supervisor (Robin Seymour) personally stamps `"N/A RS [Date]"`.
'''

# 1. Update .agents/AGENTS.md
with open('.agents/AGENTS.md', 'r', encoding='utf-8') as f:
    text_agents = f.read()

target_agents = '## SOP: Creating a New OOS Module (Template Handoff Workflow)'
if target_agents in text_agents:
    text_agents_new = text_agents.replace(target_agents, rule_0r_agents + '\n\n---\n\n' + target_agents, 1)
    with open('.agents/AGENTS.md', 'w', encoding='utf-8') as f:
        f.write(text_agents_new)
    print('AGENTS.md successfully updated with RULE 0R and RULE 0S')
else:
    print('Target not found in AGENTS.md')

# 2. Update PRD.md
rule_prd_top = '''15. **RULE 0P — Master Table Golden Standards: Robin Seymour Mandates & User Core Requirements (OOS 附表双轨终极标准规范：Robin 质保法定规范与用户核心排版铁律)**:
    - **分段横幅标准 (Canonical 4-Banner Dividers)**: `Personnel EM` (无日期)、`Biological Safety Cabinet EM for ...` (无 BSC 前缀无日期用 and)、`Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs` (无日期)、`Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs` (无日期)。
    - **严禁使用 "&" 连接符**: 一律使用英文单词 `and`。
    - **标准观察结果与注释**: `No Growth`, `Not Applicable`, `None`，异常说明首字母大写无句号。
    - **微生物双名法斜体**: 属名大写种名小写，全斜体 (*Italic*)。
    - **确证物种免染色形态 (Microbial ID Precedence)**: 有具体菌种学名时直接填报菌种，严禁再写初步革兰氏形态。
    - **多物种强制独立空行 (Mandatory Species Empty Line)**: 检出多菌种时，学名之间必须插入独立空行 (`\\n\\n`)。
    - **全表单元格字号绝对均一 (Strict Uniform Font Size - 7.0pt)**: 全表所有文字包括底层 XML 超链接统一为 7.0pt (`<w:sz w:val="14">`)。
    - **双文件交付法定归宿 (Rule 0M Compliance)**: 仅交付带表完整 PDF 与 Word 独立附表，严禁单独交付独立 Table PDF。

16. **RULE 0Q — Good Documentation Practices (GDP) N/A Sign-Off Standard (N/A + 分析员官方缩写 + 2位年份日期 DDMMMYY 规范)**:
    - 法定格式：`N/A <Analyst Initials> <DDMMMYY>`（如 `N/A QYC 21Oct26`）。
    - 严格 2 位年份（严禁 4 位），严格遵守花名册缩写（Qiyue Chen 必须为 `QYC`）。

17. **RULE 0R — Universal 3-Tier OOS Review Hierarchy & Supervisor (RS) vs QA Role Distinction (全 OOS 模块三级签名流转与主管 RS vs 质保 QA 角色法定铁律)**:
    - **Prepared by**: Qiyue Chen (QYC / 调查起草人) 签署 `Signature_0`。
    - **Reviewed by**: Robin Seymour (RS / 直接实验主管) 签署 `Signature_1`。
    - **Approved by (QA)**: QA 质保部终审批准 签署 `Signature_2`。QA 在主管审核期间尚未介入！
    - **流转顺序**: `QYC (Prepared by) ✍️` -> `RS (Reviewed by) ✍️` -> `QA (Approved by) ✍️`。

18. **RULE 0S — Universal Sterility Investigation Standards from Robin Seymour (全 OOS 模块复溶详述、设备栏 "Phase I Summary"、CR115 校准期、冲突抗辩两段式、根因与结语极简法则)**:
    - **复溶详述必写**: 冻干粉/多肽必须详细说明复溶溶媒、体积、浓度及过滤分液比例。
    - **设备栏引用**: 环境控制引用规范为 `"Phase I Summary"`（无 `"See"`），显式列出主处理 BSC 及其校准到期日。
    - **CR115 校准期**: 统一为 **`Dec 2026`**。
    - **冲突抗辩两段式**: 快速法合格 vs 药典法失败时，必须先指出 ScanRDI 无方法适用性支撑，再阐述低浓度微生物非均匀分布，证明二者不可直接对比。
    - **根因分析极简**: 聚焦复测确证同种菌判定原失败有效，严禁写冗长动力学。
    - **调查结语**: 双重判定初始与复测均为有效失败。
    - **Section E 处置**: 起草时留空，由主管 Robin 填写 `"N/A RS [Date]"`。
'''

with open('PRD.md', 'r', encoding='utf-8') as f:
    text_prd = f.read()

target_prd = '## Overview'
if target_prd in text_prd:
    text_prd_new = text_prd.replace(target_prd, rule_prd_top + '\n\n---\n\n' + target_prd, 1)
    with open('PRD.md', 'w', encoding='utf-8') as f:
        f.write(text_prd_new)
    print('PRD.md successfully updated with RULE 0P, 0Q, 0R, and 0S')
else:
    print('Target not found in PRD.md')
