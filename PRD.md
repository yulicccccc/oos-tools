# Product Requirements Document: OOS Tools

## 📌 LATEST CRITICAL UPDATES (PINNED AT TOP — READ LINES 3000 TO END FOR FULL DETAILS)
> **MANDATORY FOR ALL AGENTS**: `PRD.md` exceeds 3,200 lines (beyond a single 800-line `read_file` call). In addition to reading this top section, you **MUST call `read_file(Path="PRD.md", StartLine=3000)`** to read `Rule 61`, `Rule 62`, `Rule 70` (Robin Seymour Supervisor Review Standards), and `Rule 71` (Official L-Suite Figure 6 Layout).
> - **Supervisor Authority (`Rule 70.0`)**: **Robin Seymour (RS)** is Qiyue Chen's direct supervisor. All RS review edits are mandatory law.
> - **ScanRDI Event Count (`Rule 70.1`)**: Must match raw **Scan Data Packet** (`19 total events, 4 confirmed microbial events`), never analyst EagleTrax entries. Ask the user to verify if raw Scan Data Packet is hard to locate.
> - **Official L-Suite Layout (`Rule 70.2` & `Rule 71` / Figure 6)**:
>   - **CR145 (`E001979`, Cert Due `Nov 2026`, ISO 7)**: Section I has **`BSC E001938`** (Cal Due `Jul 2027`, Strip #6, Processing/Filtration) & **`BSC E001319`** (Strip #5); Section II has **`BSC E001317`** (Strip #7). Room Strips #1–#4.
>   - **CR144 (`E001978`, Cert Due `Nov 2026`, ISO 7)**: Section I has **`BSC E001937`** (Cal Due `Jul 2027`, Strip #13, Changeover) & **`BSC E001988`** (Strip #12); Section II has **`Incubator E001034`**. Room Strips #8–#11.
>   - **Room 143 (ISO 8)**: Gowning Area + Section I (`Scan E002017`, `Table`, `Incubator E001032`, `Scan E001511`) + Section II (`Scan E001230`, `Scan E001040` [Cal Due `Oct 2026`]).
>   - **Room 142 (ISO 8)**: Outermost entrance room (`Sink`, `Cart`).
> - **Active Air Morphology Comparison Ban (`Rule 70.4`)**: Never claim environmental cocci (e.g., *S. epidermidis*) have different morphology from cocci in the sample.
> - **RS Streamlined Phrasing (`Rule 70.5–70.6`)**: Use RS's exact monthly cleaning sentence and streamlined concurrent positive sentence (omitting third-party client/product/analyst names).
> - **GDP 2-Digit Year N/A Sign-Off (`Rule 62`)**: Strictly `N/A QYC DDMMMYY` (e.g., `N/A QYC 21Oct26`).








## 🚨🚨 TOP-PRIORITY LOCKED MANDATES (RULE 0A & 0B — READ FIRST)



1. **RULE 0A — Analyst Identity Roster, Zero Hallucination & Mandatory Ask-First Firewall**:



   - **Root Cause**: EagleTrax usernames are `[First Initial] + [Last Name]` (`ebennett`, `enioupin`, `mjang`, `dkansara`, `jochoa`, `aodugbesi`). **NEVER expand the first initial by guessing a common first name!** There is **NO Elizabeth Bennett**, **NO Eniola Nioupin**, and **NO Min Jang** in the laboratory!



   - **Mandatory Roster**: 严格收录 33 位在册无菌实验室分析员与历史分析员唯一法定身份（详见 Rule 46 权威总表与 `utils.py`）。



     - 🚨 **核心姓名合规绝对锁定**：



       - `GL` / `gli` $\to$ **Guanchen Li (GL)**（严禁使用口语化昵称 "David Li"！）



       - `KSM` / `ksilva` $\to$ **Karla Silva (KSM)**（严格使用 Karla Silva，严禁使用母姓 "Karla Silva Machuca"！）



       - `MRB` / `mbythatagari` $\to$ **Muralidhar Bythatagari (MRB)**（严格使用 Muralidhar Bythatagari，严禁带中间名 "Reddy"！）



       - `MN` / `mnaeem` $\to$ **Maryam Naeem (MN)**（官方确认 2 位缩写为 MN！）



       - `NT` / `ntrinh` $\to$ **Nicholas Trinh (NT)**（官方确认 2 位缩写为 NT，Technician 无班次）



       - `SAA` / `sallen` $\to$ **Sasha Allen (SAA)**（🚨 官方确认缩写为 SAA，严禁写成 SA!）



       - `ebennett` / `ELB` $\to$ **Ellen Bennett (ELB)**（严禁脑补为 "Elizabeth Bennett"！）



       - `enioupin` / `EN` $\to$ **Elysse Nioupin (EN)**（严禁写成 "Eniola" 或发明 "ENI" 缩写！）



       - `mjang` / `MJ` $\to$ **Mukyung Jang (MJ)**（严禁写成 "Min Jang"！）



       - 🚨 **Soniyah Reed 从未出现在本实验室，坚决不予收录！**



   - **Mandatory Ask-First Rule (严禁自作聪明·不懂必问)**: Whenever encountering an unknown username, missing logbook, or uncertain detail, **STOP immediately and ask the user** — NEVER guess!



2. **RULE 0B — Robin Seymour (RS) Table 2 Section Divider Standard**:



   - Banner 1: **`Personnel EM`** (Canonical: NO date suffix! In single-day bracketing where QA requested, `Personnel EM for DDMMMYY` is accepted.)



   - Banner 2: **`Biological Safety Cabinet EM for E00XXXX and E00YYYY`** (Canonical: NO `(BSC)`, NO date suffix! In single-day bracketing, `(BSC) Biological Safety Cabinet EM for BSC E00XXXX and BSC E00YYYY` is accepted.)



   - Banner 3 & 4: **`Weekly Active Air / Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs`** (Must list all room numbers, NO date suffix! Robin Seymour streamlined `(E001979, L-Suite)` to simply `(L-Suite)`.)



3. **RULE 0C — Standard Microbiology SOP Designations (MICRO-SOP-2)**:



   - In all OOS reports (ScanRDI, Celsis, USP <71>, and EM), environmental monitoring plate incubation and facility monitoring MUST strictly cite **`MICRO-SOP-2`** (`Environmental Monitoring of the Clean-room Facility`), NEVER `SOP 2.600.002`.



4. **RULE 0D — Zero Guessing on Positive Controls, Reagent Lots & Test Bench Data (阳性对照与核心实验台账零脑补·不明必问铁律)**:



   - **cGMP 法定证据链铁律**: 阳性对照菌株（Standard / Positive Control Organism）、批号（Lot #）、效期（Expiration Date）属于质保（QA/Robin Seymour）审核与合规归档的法定关键数据。**严禁从旧案卷拷贝，严禁使用模板默认占位符，严禁擅自脑补！**



   - **强制停步提问门禁**: 凡在 EagleTrax 系统电子记录或现场纸质扫描件中未查明阳性对照菌种及批号效期时，**必须立即停步，向用户主动提问获取**（例如直接向用户询问：“请提供本次检验的阳性对照菌种、Lot 批号与效期”），绝不允许擅自填入虚假或猜测数据导致用户被质保批评！



   - **关键实验参数零脑补延伸**: 包含生物安全柜日常消毒记录批号、双氧水指示卡/IPA 批号、培养基批号效期、设备运行 Session/Record 编号等，凡无法 100% 确认真实性的，一律先向用户确认，宁可多问一句，绝不乱写一字！



5. **RULE 0E — Strict Prohibition of "&" in Official Report Forms & Dividers (严禁在正式报告表单与标题中使用 "&" 连接符)**:



   - **多套间/多洁净室分行规范**: 在 Cleanroom 认证栏目（如 Page 2 `Text Field32`/`33` 及 Word 表格 `smart_cr_id`），多套间严禁使用 `&` 并列（严禁写成 `E001738 (Suite 116) & E001978 (L-Suite)`）。必须采用标准多行对齐：例如 L-Suite 双房间填报第一行 `E001978 (CR144)`、第二行 `E001979 (CR145)`，右侧对应校准期 `Nov 2026\rNov 2026`（字号 5.0pt）；若涉及 CR116 (`E001738`) 则对应 `Dec 2026`。



   - **正文与标题规范**: 在段落叙述和分段横幅标题中，一律使用标准英文单词 **`and`**（如 `BSC E001312 and E001937`），绝不使用 `&`。杜绝给人以非正式、粗糙或排版怪异的印象。



6. **RULE 0F — Strict Prohibition of Absolute Claims ("100%") in Narrative (严禁在调查报告中使用 "100%" 绝对化用语)**:



   - **cGMP 质保铁律**: 在无菌检验与环境监测 OOS 调查报告叙述中，**严禁使用 "100% absence of microbial recovery" 或 "100% sterile" 等绝对化词汇**。质保（QA/Robin Seymour）审核认为此类表述过于绝对。



   - **法定标准用语**: 必须统一规范写作：



     `"Upon review of the environmental monitoring data, absence of microbial recovery (No Growth) was demonstrated across all critical ISO 5 zones and personnel monitoring on the date of testing."`



7. **RULE 0G — Standard Phrasing for 6-Month Analyte History with Prior Failures & Zero Failures (6 个月特定分析物历史、追踪台账溯源与句式铁律)**:



   - **官方受控追踪台账数据源 (Robin Seymour SharePoint Ground Truth)**:



     - 6 个月历史核查必须严格基于官方 SharePoint 受控台账：



       `https://eagleanalytical-my.sharepoint.com/:x:/r/personal/rseymour_eagleanalytical_com/_layouts/15/Doc.aspx?sourcedoc=%7B8A575D4F-67AB-4253-B1E2-08A540577BCB%7D&file=Sterile%20Lab%20-%20OOS%20Tracking%20Log.xlsx&fromShare=true&action=default&mobileredirect=true`



     - **强制问询门禁**: 若无法访问上述台账或不确定历史记录，**必须立即向用户主动提问获取**，严禁擅自断言！



   - **Case A: 无前科句式 (Zero Prior Failures)**:



     - 必须明确包含无菌检验测试方法后缀（如 `under ScanRDI sterility testing.` / `under Celsis sterility testing.` / `under USP <71> sterility testing.`）：



       `"Analyzing a 6-month sample history for [Client Name], this specific analyte "[Analyte Name]" has had no prior Out-of-Specification (OOS) occurrences under [Test Method] sterility testing."`



   - **Case B: 存在前科三步曲句式 (Prior Failures)**:



     - **基础陈述句**: `Analyzing a 6-month sample history for [Client Name], this specific analyte "[Analyte Name]" has had [one prior failure / N prior failures], documented under [OOS-ID] for submission [ETX-ID], tested on [DDMMMYY].`



     - **形态学差异对比句**: `The microorganism detected in the previous investigation exhibited [clustered circular-shaped / previous morphology] morphology, which differs from the [current morphology] observed in the current investigation.`



     - **前案调查回顾与闭环评价句**: `The previous failure investigation was reviewed and no laboratory error or assignable laboratory cause was identified.`



   - **语法与标点铁律**: 送检单号 `[ETX-ID]` 之后**必须使用逗号隔开**（`, tested on [DDMMMYY].`），杜绝介词堆叠（`under... for... on...`）。



8. **RULE 0H — Cleanroom Plural Conjunction Standard (同级多洁净室并列使用 "and" 连词铁律)**:



   - 在描述周检浮游菌或表面监测时，并列的多间同级别洁净室必须使用连词 **`and`**，严禁使用 `or`（例如：必须写成 `in the ISO 7 cleanrooms (145 and 144)`，绝不可写成 `(145 or 144)`），以准确反映所有对应房间均实施了采样且均无微生物生长。



9. **RULE 0I — ISO 8 Ambient Exposure Technical Defensibility (ISO 8 暴露技术防守用语规范)**:



   - 严禁在叙述中使用绝对化无接触表述（如 "The test sample does not come into contact with ambient ISO 8 air"）。



   - 统一使用质保认可的防御性用语：



     `"There is minimal exposure of the test sample to ambient ISO 8 air, as samples and supplies are transferred in disinfected, closed containers on carts through the layered cleanroom suites."`



10. **RULE 0J — Multi-Suite Monthly Cleaning Historical Granularity (多套间月度清洁日期精确拆分法则)**:



    - 涉及多套间（如 Processing 在 CR116，Changeover 在 L-Suite）时，严禁合并为单一日期，必须分别列出各自套间的真实清洁日期：



      `"Monthly cleaning and disinfection of the cleanrooms (ISO 7) and their containing biosafety cabinets (ISO 5) were performed on the [DDMMMYY] for CR116 and [DDMMMYY] for L-Suite, as per MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. It was documented that all chemical indicators passed, verifying that environmental sanitization remained validated and in control prior to testing."`



11. **RULE 0K — Equipment Description Room Specificity (设备与洁净室编号精准度规范)**:



    - 在 Page 2 设备信息栏 (`Text Field32`/`33`) 与 Word 模板 `smart_cr_id` 中，统一使用具体房间代号：



      - L-Suite 双房间标准（Processing in CR145 + Changeover in CR144）：第一行 `E001978 (CR144)`，第二行 `E001979 (CR145)`，对应校准期均为 `Nov 2026\rNov 2026`（5.0pt）

      - 跨套间标准（如 CR116 + CR144）：`E001738 (CR116)`（校准期 `Dec 2026`）与 `E001978 (CR144)`（校准期 `Nov 2026`）



12. **RULE 0L — Cross-Contamination Opening, Multi-Shift Disclosure & Concluding Standard Phrasing (交叉污染全天批次评估、多班次/多分析员透明披露与结论质保法定用语)**:



    - **交叉污染段落标准起始句**: `"All samples processed on the same day were also evaluated for evidence of sample-to-sample cross-contamination."`



    - **同日其他阳性样本透明披露与隔离证明法则 (Multi-Shift / Multi-Analyst Disclosure Rule)**:



      - 严禁在未核实全天所有班次、所有分析员的情况下草率断言"其他所有人员检验结果均为阴性"。



      - 若当日存在其他阳性样本（即使发生在不同班次、不同分析员、不同套间或不同 BSC）：



        1. **主动透明披露**: 详述该阳性样本案卷号（如 `OOS-XXXXXX`）、送检单号（`ETX-XXXXXX-XXXX`）、客户、药品规格及运行序位。



        2. **极简高效隔离证明法则 (Streamlined Segregation Phrasing - 聚焦不同人员与不同时间)**:
          - 🚨 **用户定稿极简原则**: 无需堆砌过多冗余琐碎信息（如无需专门交代 Prepper 姓名或强调外围套间层级名称）。
          - 核心证明聚焦三大要素：**班次（Shift）、检验处理分析员（Processing Analyst / 样本处理序号）、主操作设备（Primary Containment Equipment / BSC）**。
          - **法定标准精炼句式**:
            `"However, that sample was processed by [Analyst] ([N]th sample processed) during the [Shift]. The two testing events were segregated by shift, processing analyst and primary containment equipment."`



        3. **手套消毒与无菌操作合规重申**: 说明操作人员按 SOP 在样本间严格实施手套消毒与无菌规程。



        4. **排除交叉污染与共用媒介结论**: 证实不存在样本间交叉污染或共同实验室污染源。



    - **调查结论标准句**: `"Therefore, the likelihood of laboratory error contributing to this failure is considered minimal, and the original result is deemed valid."`



13. **RULE 0M — Official 2-File Deliverable Protocol & Canonical Desktop Destination (最终交付物双文件法定铁律与企业 OneDrive 桌面唯一落盘路径铁律)**:

    - **法定桌面落盘路径铁律 (Canonical Desktop Destination Rule)**:
      - 凡是用户或流程要求将 OOS 交付物保存至“桌面（Desktop）”时，**必须严格唯一定位并保存至企业 OneDrive 桌面物理路径**：
        `C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\`
      - 🚨 **严禁仅保存至本地物理桌面** `C:\Users\qchen\Desktop`，彻底杜绝多路径引发的“找不到文件”或版本不同步问题！

    - **用户法定指令 (User Golden Mandate)**: 凡完成任何 OOS 调查报告生成或修订任务，**严格只向用户输出并提供以下 2 个核心文件**，严禁输出任何冗余多余文件：

      1. **文件 1 (带 Table 的完整 PDF 版主案卷 — 官方转录透视命名规范)**:

         - **转录官方表单正式命名**: 凡转录至 ZenQMS 官方空白母表之正式案卷，文件名**必须统一冠以表单编号与同日下载日期前缀**：

           - Phase I: CORP-FORM-21- <DD MMM YYYY> OOS-<ID> <Client> - <Test>.pdf

             （例如：CORP-FORM-21- 01 Oct 2026 OOS-262289 Link Pharmacy (E19112) - ScanRDI.pdf）

           - Phase II: CORP-FORM-22- <DD MMM YYYY> OOS-<ID> <Client> - <Test>.pdf

             （例如：CORP-FORM-22- 10 Sep 2026 OOS-262098 Solyn LLC (E75000) - Phase II.pdf）

         - **核心目的 (一眼核验)**: 使审查人与用户能够**一眼直观核验该案卷是否已原生完成官方表单转录**，以及其采用的母表下载日期是否与阳性发现日期 100% 吻合！

         - 内容包含 6 页 AcroForm 官方调查表单及末尾附加 Page 7/8 Table 1 & Table 2 附表的完整合并版 PDF，严格排版无溢出）。

      2. **文件 2 (Word 版独立附表)**: `Tables OOS-<ID> <Client> - <Test>.docx`

         （包含 Table 1 样品信息与 Table 2 环境监测数据的独立 Word 附表文档）。

    - **绝对严禁交付的冗余文件 (Strictly Prohibited Deliverables)**:

      - 🚨 **严禁提供** `Tables OOS-<ID> <Client> - <Test>.pdf`（独立 Table PDF）。

      - 🚨 **严禁提供** `OOS-<ID> <Client> - <Test>.docx`（主表单 Word 版）。

      - 彻底杜绝多余文件造成混淆。

    - **双路径呈现规范**: 对上述 2 个文件，必须同时提供 Markdown 可点击超链接与 Windows 原生绝对路径（方便复制粘贴直接打开）。









14. **RULE 0N — Strict Prohibition of Arbitrary Font Size Alteration & Formatting Tampering during Copy-Paste (严禁在复制粘贴/文本与表格填充时擅自改动任何字号大小与排版格式铁律)**:

    - **绝对零擅自改动字号与格式**: 凡用户手动修改、确认或粘贴的内容，**严禁自行更改其字体大小（Font Size）、字体名称（Font Family）、字形（Regular/Italic/Bold）、行间距或对齐方式**！

    - **同类数据字号严格统一**: 表格内部各单元格（如 Table 1 样品信息、Table 2 环境监测、Microbial ID、Observation、Notes 等），同类行或同一数据块内的字体大小必须**严格保持完全统一**（例如统一 7.0pt 或 7.5pt），严禁个别行或个别单元格擅自变大变小导致视觉突兀！

    - **物种空行与结构换行严格保留**: 当多菌种（Microbial ID）或多条目按用户要求在物种之间保留空行以增强可读性时，复制粘贴或代码拼接时**必须完整保留原样空行**，绝对严禁擅自压缩、删除空行或合并为单行！

    - **杜绝自作主张的“顺手优化”**: 用户已定稿的排版属于神圣资产。严禁打着“让版面更好看”或“自动适应”的旗号修改字号。





15. **RULE 0P — Master Table Golden Standards: Robin Seymour Mandates & User Core Requirements (OOS 附表双轨终极标准规范：Robin 质保法定规范与用户核心排版铁律)**:
    - **分段横幅标准 (Canonical 4-Banner Dividers)**: `Personnel EM` (无日期)、`Biological Safety Cabinet EM for ...` (无 BSC 前缀无日期用 and)、`Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs` (无日期)、`Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs` (无日期)。
    - **严禁使用 "&" 连接符**: 一律使用英文单词 `and`。
    - **标准观察结果与注释**: `No Growth`, `Not Applicable`, `None`，异常说明首字母大写无句号。
    - **微生物双名法斜体**: 属名大写种名小写，全斜体 (*Italic*)。
    - **确证物种免染色形态 (Microbial ID Precedence)**: 有具体菌种学名时直接填报菌种，严禁再写初步革兰氏形态。
    - **多物种强制独立空行 (Mandatory Species Empty Line)**: 检出多菌种时，学名之间必须插入独立空行 (`\n\n`)。
    - **全表单元格字号绝对均一 (Strict Uniform Font Size - 7.0pt)**: 全表所有文字包括底层 XML 超链接统一为 7.0pt (`<w:sz w:val="14">`)。
    - **双文件交付法定归宿 (Rule 0M Compliance)**: 仅交付带表完整 PDF 与 Word 独立附表，严禁单独交付独立 Table PDF。

16. **RULE 0Q — Good Documentation Practices (GDP) N/A Sign-Off Standard (N/A + 分析员官方缩写 + 2位年份日期 DDMMMYY 规范)**:
    - 法定格式：`N/A <Analyst Initials> <DDMMMYY>`（如 `N/A QYC 21Oct26`）。
    - 严格 2 位年份（严禁 4 位），严格遵守花名册缩写（Qiyue Chen 必须为 `QYC`）。
    - **签署日期铁律 (Date of Drafting/Signing Ground Truth)**：N/A 后的签署日期必须严格为**撰写/签署当天的物理真实日期**（Date of Drafting/Signing，例如今天处理即为 22Oct26），**绝对严禁**误填为事故发生日（Date of Incident）或实验检验日（Test Date）！

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


---

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



    *   **SOP Reference:** MICRO-SOP-2 (Environmental Monitoring of the Clean-room Facility)



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



        *   **Product Contamination Root Cause Conclusive Confirmation:** Initial test under `ETX-260902-0505` failed on Day 6 in TSB (*Microbacterium sp. PM5*, `ETX-260910-0290`, Gram (+) rods, Plate A TSA & Plate B SDA). Retest under `ETX-260914-0470` (prepped by Andrew Carrillo, processed by Ayomide Odugbesi) replicated the exact failure kinetics, turning turbid positive on Day 6 in TSB, with sequencing under `ETX-260921-0498` confirming the identical strain: ***Microbacterium sp. PM5*** (`Gram (+) short rods`, Plate C TSA & Plate D SDA). Definitively refuted preliminary analyst handling error hypothesis during reconstitution and confirmed inherent batch contamination (Lot `2608-216`).



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



    41. *Universal Microbial Binomial Nomenclature Italicization Standard & End-to-End Autonomous Table 1 & 2 Generation Across All OOS Modules (全 OOS 模块微生物拉丁学名绝对强制斜体与全流程台账表格生成规范 - 🚨 永久铁律):*



        *   **1. 拉丁双名法微生物名称绝对强制斜体 (Mandatory Binomial Italicization):**



            - 在所有 OOS 调查报告的表格（Table 1, Table 2, Table 3）及正文叙述中，凡出现微生物属名和种名（如 *Micrococcus luteus*、*Staphylococcus capitis*、*Corynebacterium ureicelerivorans*、*Moraxella osloensis*、*Kocuria rhizophila*、*Curvularia sorghina*、*Staphylococcus hominis*、*Staphylococcus saprophyticus*、*Scolecobasidium mirabile* 等），**必须严格设置为斜体 (`italic = True` / `*Genus species*`)**！



            - 染色描述与形态学分类词（如 `Gram (+) short rods`、`Gram (+) cocci`、`Gram (-) coccobacilli`、`Hyphae`、`sp.`、`ATCC 6931` 等）保持正体。



        *   **2. 全模块端到端独立 Table 1 & Table 2 自动化生成生态 (Universal Standalone Tables Capability):**



            - 覆盖所有 OOS 模块（`ScanRDI`、`Celsis`、`USP <71>`、`EM`），AI 具备穿透 EagleTrax 自动化查询、清点 6 大纸质台账扫描件、提取活菌染色与鉴定、并直接生成完美适配**单页（Page 1 of 1）**独立 Word (`.docx`) 及 PDF 表格交付物的完整闭环能力。



            - 洁净室周检多点检出（如 L-Suite 3 处检出）严格遵循单行合并单元格规范，通过精准 `space_before` 调优，实现检出点、原生交互超链接与斜体菌株列表在横向基线上的绝对 1:1 对齐。



    42. *OOS Investigation Lifecycle Role Separation & Drafting vs. Review Phase Boundary Standard (调查生命周期角色分离规范：撰写阶段绝对严禁越权填写 Reviewer / Closure 区域 - 🚨 永久铁律):*



        *   **🚨 角色权限红线背景 (Drafter vs. Reviewer Strict Boundary):** 在 cGMP 质量体系与 ZenQMS 官方调查表单（如 CORP-FORM-21 / 3.100.019.F01）流转中，AI 与调查员在草拟阶段的身份仅为 **“调查撰写人 / 调查员 (Drafter / Investigator)”**，绝对不是审核人或批准人（“我们只是写的人，并没有人 review 了哦”）。严禁越权代填后续审批流专用的审核意见与结案签字！



        *   **1. 表单第 6 页（Page 6）草拟阶段严格留白规范 (Mandatory Page 6 Blank Fields for Drafting Mode):**



            - **审核评语栏留白 (`Text Field52: Reviewer comment(s)`):** **必须绝对留空 (`""`)**！严禁预填任何总结性审核评语（如 “The investigation was reviewed and found complete...”），该栏目必须由质保审核人（QA Reviewer / Lab Manager）在正式审批时手填或电子签署。



            - **质量处置与结案勾选留白 (`Section E Quality Disposition`):** `Check Box90`（Phase I Investigation Closed）、`Check Box91`（Phase I Investigation Not Closed）、`Check Box92`（Phase II Not Required）**必须全部保持未勾选 (`Off`)**！结案处置判定属于管理层审批职权，草拟阶段严禁擅自提前关闭。



            - **审批签字人打印栏留白 (`Text Field54` & `Text Field55`):** `Prepared by (Print)` (`Text Field53`) 填写实际撰写人（如 `Qiyue Chen`）；而 `Lab Manager/Designee (Print)` (`Text Field54`) 和 `Approved by QA Representative (Print)` (`Text Field55`) **必须严格留空 (`""`)**。



            - **调查员调查结论勾选 (`Section D`):** 调查员正常勾选 `No, the OOS result is valid.` (`Check Box88: Yes`, `Check Box87: Off`, `Check Box89: Off`)，确认未发现实验失误，原始阳性结果有效。



        *   **2. ZenQMS 官方下载表单（CORP-FORM-21）转录保护与无缝缝合规范 (Official ZenQMS Transcription Standard):**



            - **强制版本备份优先原则:** 对用户提供的任何来自 ZenQMS 的官方下载模板（如桌面 `CORP-FORM-21 - P1 - 16 Sep 2026.pdf`），在执行自动化填充前，必须首先在 `.history/` 目录下生成带时间戳的安全副本备份（如 `.history/CORP-FORM-21 - P1 - 16 Sep 2026_backup_YYYYMMDD_HHMMSS.pdf`），严禁直接裸写原文件。



            - **系统审计底纹与原生页脚绝对保留:** 严禁破坏 ZenQMS 导出的官方审计底脚（如 `Printed by rseymour@eagleanalytical.com from https://app.zenqms.com on 16 Sep 2026 at 19:15 UTC`）。



            - **表单特定差异化字段精准适配:** 针对 ZenQMS 空白表单特有的可填文本框（如 Page 1 `Text Field1: Scan RDI Sterility Test`），必须精准补全；对易截断的紧凑单元格（如 Page 2 `Text Field32/33` 耗材批号效期），应用精准自适应字号（4.4pt），彻底杜绝 Acrobat `+` 号溢出截断。



            - **Page 7 台账表格无缝缝合:** 将已生成并通过严谨审计的单页独立表格（Table 1 & Table 2）高保真转换为 PDF，并自动作为 Page 7 无缝缝合至 6 页官方主表之后，形成合规完备的 7 页调查卷宗，随时供用户上传至 ZenQMS。



    43. *Sterility Investigation Phase 1 vs. Phase 2 Scope Separation Architecture (无菌检测 Phase 1 与 Phase 2 调查范围绝对解耦架构 - 🚨 黄金标杆规范):*



        *   **🚨 核心业务逻辑与样本数区分机制 (Only Phase I vs. With Phase II Distinction - 根治混淆终极铁律):**



            - **为什么有的 OOS 报告 Table 1 只有 1 个样本？** 因为那些案子（如日常常规 OOS 或实验室原因关闭的案子）**只有 Phase I**，根本没有发生复测（No Retest / No Phase II），所以物理上只存在 1 个样本！



            - **为什么有 Phase II 的案子 Table 1 必须包含 2 个样本？** 凡是一起 OOS 调查开启并存在 Phase II（即发生了 Retest 复测，如 `OOS-262098` 与 `OOS-261987`），其调查母表（Master Tables）**必须始终统筹呈现两阶段的两个测试样本**（Row 1: 初始测试样本；Row 2: 复测测试样本）！



            - **严禁混淆：** 绝不能看到“其他只有 Phase I 的案子表格只有 1 个样本”，就盲目将“有 Phase II 的案子”的 Phase I 附件表格强行删减成单样本！两类案子的业务场景截然不同！



        *   **🚨 终极架构原则与标杆对照 (Reference: OOS-261987 Uriel Pharmacy P1 & P2 Signed by OA-RS):** 当一起无菌检测 OOS 事件涉及初始测试（Original Test）与复测（Retest）两个阶段时（如 `OOS-262098` / `OOS-261987`），调查案卷必须遵循严格的范围解耦架构：



        *   **1. 独立表格附件 (Master Tables DOCX/PDF):**



            - 维持对两阶段测试的**统筹全量呈现**：Table 1 必须同时列出初始样品与复测样品信息（双样品横向对齐）；Table 2 呈现初始测试环境监测；Table 3 呈现复测环境监测。用户定论：“table是对的，是需要2个tests （origianl和retest的信息）”。



        *   **2. Phase 1 官方调查报告 (CORP-FORM-21 / 3.100.019.F01):**



            - **仅针对初始测试（Original Test ONLY）！**



            - Page 1：样品 ID 仅填初始样品（如 `ETX-260902-0505`）；测试日期仅填初始测试日期（如 `04-Sep-2026`）；分析员仅填初始人员（Prepper / Processor: Alex Saravia, Reader: Elysse Nioupin）；事件描述仅描述初始阳性；Section B 仅针对初始测试记录。



            - Page 2：设备与洁净室仅记录初始测试使用的 BSC（如 `ISO 5 BSC E001316`）与洁净区（`CR114`）。



            - Page 3..Page 5 叙述正文：**严禁提前提及复测、严禁出现复测分析员与复测样品号！** 严格遵循 `OOS-261987 P1` 叙述范式，仅针对初始测试的实验操作、培养读数、菌株初步鉴定、初始环境监测防守闭环进行调查，最终判定初始阶段未发现实验室人为/环境偏差，初始阳性有效（Page 6 勾选 `Check Box88` "No, the OOS result is valid"）。



        *   **3. Phase 2 官方调查报告 (CORP-FORM-22 / 3.100.019.F02):**



            - **承接并详尽展开复测调查（Retest Investigation）！**



            - Page 1：标明初始样品与复测样品号、复测日期（`15-Sep-2026`）、复测分析员（`Ayomide Odugbesi`）、复测结果（`Fail`）。



            - Page 3：概括性回顾初始测试调查历程（Recap of Original Investigation）。



            - Page 4：以清晰小标题 **`RETEST UNDER SUBMISSION ETX-XXXXXXXX-XXXX.`** 正式展开复测的完整叙述（复测操作人员、Suite 115 / BSC E001314 环境、培养第6天浑浊重现、测序同源菌株 *Microbacterium sp. PM5* 确证、前期 ScanRDI 合格原因技术论证等）。



            - Page 5：归结根本原因（产品固有污染 / 外部因素 External Phenomena）并签署关闭。



    44. *EagleTrax Benchmark Extraction & Bench Operator Ground Truth Alignment (EagleTrax 原始凭证穿透与实验台账真实操作人校准规范 - 🚨 永久铁律):*



        *   **1. 实验台账真实操作人穿透提取 (Raw Bench Operator Penetration):**



            - 在无菌检测调查中，前置准备（Sample Prep）、过滤接种（Sample Analysis / Processing）与培养读数（Daily Reads）往往由不同人员分工操作。



            - 必须严格穿透 EagleTrax `#EventHistory` 与台账记录核查真实人员与时间戳：如 `ETX-260902-0505` 中，前置准备由 `isharma`（**Ishita Sharma, ISS**）于 03Sep26 20:48 完成，主处理由 `asaravia`（**Alex Saravia, ES**）于 04Sep26 22:33 完成，读数由 `enioupin`（**Elysse Nioupin, EN**）与 `acarrillo`（**Andrew Carrillo, AC**）执行。



            - 严禁将处理分析员（Alex Saravia）盲目套用为前置准备人。在表单与叙述中必须完整体现：Prepping Analyst: Ishita Sharma (ISS)；Processing Analyst: Alex Saravia (ES)；Reading Analyst: Elysse Nioupin (EN)。



        *   **2. 样品取样量与过滤介质真实规格校准 (Sample Quantity & Membrane Filtration Alignment):**



            - 严格穿透 Worksheets 提取过滤总量与容器规格：如 `45 gm added per media` across `3 vials`（每瓶 30 mL/gm，共 3 瓶 90 gm，分别过滤 45 gm 至 FTM 与 45 gm 至 TSB），必须在 Section A、叙述第 8 段及材料说明中精准表述，杜绝笼统或模糊数字。



        *   **3. 耗材批号与系统附件无缝闭环 (Consumables & System Documents Audit):**



            - 真实耗材（FTM Lot 2567320 Exp 12/16/2026, TSB Lot 2548010 Exp 04/19/2027, Fluid D Lot 685561 Exp 10/31/2026）必须在 EagleTrax 耗材系统与调查卷宗中保持绝对一致。



    45. *Rapid ScanRDI Negative vs. Multi-Vial USP <71> Positive Discrepancy Reconciliation Standard (ScanRDI 快速检测阴性与 USP <71> 双重阳性矛盾之科学归因与合规抗辩规范 - 🚨 终极质保规范):*



        *   **🚨 现象本质与科学常识 (Microbiological Ground Truth):** 在无菌制剂批次检验中，看似“离谱”的矛盾（同批次前期 ScanRDI 判定 Pass，而后续送检的初始 USP <71> 与复测 USP <71> 双双判 Fail），在微生物学与 cGMP 调查中具有完全自洽的科学解释：这是无菌药品批次中**低浓度、非均匀分布微生物污染（Heterogeneous, Low-Concentration Bioburden / Discrete Contamination Pockets）**的典型特征！



        *   **三大合规抗辩支柱 (Three Evidentiary Pillars in Investigation Narrative):**



            1. **严重抽样不对称性 ($n = 1$ vs $n = 7$ 瓶取样量悬殊):** 



               - 穿透 EagleTrax 原始凭证（`3.pdf`）：ScanRDI 检测（`ETX-260807-0602`）送检及测试的小瓶数仅为 **1 瓶（`Number of Items in Batch: 1`）**。在极微量非均匀污染的批次中，单瓶抽样有极高概率恰好抽中未受污染的阴性小瓶（0 CFU 活菌事件），因此 Scan 读数 `14/0` 正常判定为合格。



               - 相比之下，USP <71> 初始测试（`ETX-260902-0505`）抽检了 **3 瓶（90 g）**，复测（`ETX-260914-0470`）抽检了 **4 瓶（120 g）**，两次合计测试了 **7 瓶（210 g）**，极大地提高了捕获小概率污染小瓶的统计学检测灵敏度与样品代表性。



            2. **方法适用性档案缺失 (`No Method Suitability on file`):**



               - `3.pdf` 系统明确标注 `Method Suitability: No Method Suitability on file`。ScanRDI FIFU 荧光法未对该复方多肽制剂（Cagrilinitide 5mg/vial, Retatrutide 25mg/vial）进行基质相容性、滤膜洗脱效率与活菌荧光标记回收率的正式验证。



            3. **瞬时细胞计数 vs. 14 天肉汤指数级增殖放大机理差异:**



               - ScanRDI 依赖滤膜上细胞微小酶活性的即时荧光检测，对代谢极度迟缓或损伤的微量菌体可能存在检出限阈值。



               - USP <71> 是 14 天长周期 100 mL 营养肉汤培养（TSB/FTM），即使样品中仅含有 1 个具有活力的休眠态菌体，在连续恒温培养下也会呈对数期几何级数增殖，最终在第 6 天爆发宏观肉眼浑浊。更关键的是，两次独立的 USP <71> 培养物经过 Microgen / 16S rRNA 测序均精准锁定了**完全相同的污染菌种 *Microbacterium sp. PM5***。



        *   **调查结论合流与反证效应:**



            - 这一前期 ScanRDI 阴性结果不仅不削弱 OOS 调查，反而提供了关键反证：强有力地证明了批次污染并非高浓度全批次弥散，而是离散低载量分布；同时也排除了生产和取样过程中的宏观灾难性污染，与后续两次独立实验室人员/洁净室测试高度复现的 *Microbacterium sp. PM5* 形成闭环，最终确证根本原因为制剂产品本身的低水平非均匀固有污染。



    46. *Microbiology Analyst Ground Truth Identity & Tri-Module Anti-Hallucination Standard (全部 3 种无菌 OOS 模块 [ScanRDI, Celsis, USP <71>] 分析员真实法定姓名与官方缩写绝对守恒铁律 - 🚨 终极防线):*



        *   **🚨 终极合规背景与幻觉根本病灶剖析 (The EagleTrax Root Cause & Zero Tolerance):** 



            - **根本病灶剖析 (The Root Cause)**: 在 EagleTrax 网页系统中，所有分析员账号格式均为统一的 **`[名字首字母 First Initial] + [法定姓氏 Last Name]`**（例如 `ebennett` = e + bennett, `enioupin` = e + nioupin, `mjang` = m + jang, `dkansara` = d + kansara, `jochoa` = j + ochoa, `aodugbesi` = a + odugbesi）。



            - 过去 AI 之所以屡犯荒唐的幻觉（把 `ebennett` 臆想成 Elizabeth Bennett、把 `enioupin` 臆想成 Eniola Nioupin、把 `mjang` 臆想成 Min Jang），**其唯一祸根就是在读取到网页系统账号时，擅自用大模型臆测补全首字母，而没有绝对、严格地穿透唯一官方权威花名册查询！**



            - **🚨 终极不可逾越铁律**:



                1. 实验室**根本没有 Elizabeth Bennett，只有 Ellen Bennett (ELB)**！



                2. 实验室**根本没有 Eniola Nioupin，只有 Elysse Nioupin (EN)**！



                3. 实验室**根本没有 Min Jang，只有 Mukyung Jang (MJ)**！



                4. **绝对严禁对任何 EagleTrax 账号的首字母进行“脑补展开”！所有账号必须 100% 查表映射，表中未收录者必须立即向用户停步请示，严禁私自猜测！**



        *   **1. 全部 3 大无菌 OOS 模块（ScanRDI, Celsis, USP <71>）统一适用架构 (Tri-Module Universal Invariance):**



            - **全模块覆盖**: 无论当前调查属于 **ScanRDI**（快速激光荧光法）、**Celsis**（ATP 生物发光法）还是 **USP <71>**（经典薄膜过滤与直接接种培养法），人员解析引擎与命名规范必须 100% 保持完全一致，绝不允许因测试类型不同而出现割裂或单模块私自实现。



            - **全环节穿透**: 样品前置准备人（Prepping Analyst）、过滤接种处理人（Processing Analyst）、换台换机人（Changeover Analyst）、培养读数人（Reading Analyst）及转接人（Subculture Analyst）必须全量经由官方权威词典解析。



        *   **2. 官方权威身份、缩写与班次权威总表 (Official Ground Truth Roster & Shift Mapping):**



            - 🚨 **三大铁律法则**：



              1. **唯一法定真名铁律**：所有人名必须 100% 遵照 EagleTrax/cGMP 系统的官方真实全名！



                 - `GL` 必须且仅能写作 **`Guanchen Li`**（🚨 绝对严禁写成口语化昵称 "David Li"！）。



                 - `KSM` 遵照用户明确指示，严格写作 **`Karla Silva`**（🚨 严禁使用带母姓的 "Karla Silva Machuca"！）。



                 - `MRB` 遵照用户明确指示，严格写作 **`Muralidhar Bythatagari`**（🚨 严禁带中间名 "Reddy"！）。



                 - `CGS` 严格写作 **`Clea S. Garza`**（系统真名，严禁使用 "Santos"）。



                 - `GS` 严格写作 **`Gabrielle Surber`**（系统真名，严禁使用 "Gabby"）。



              2. **一人一码**：每位分析员有且仅有一个法定唯一 Initials，绝对不存在任何“备用”或“兼容”缩写！



              3. **班次规则**：Technician **无班次（统一标注 N/A）**；Microbiologist I/II 严格对应 10 月官方排班表（1st: 5am-2pm / 2nd: 9am-6pm / 3rd: 1pm-10pm）。



              4. 🚨 **Soniyah Reed 绝对排除**：用户校对核实其从未在无菌实验室出现过，严禁收录入分析员花名册！







            | # | 唯一法定 Initials | 英文全名 (Full Name - 法定真名) | EagleTrax 账号 | 角色 / 岗位 (Role) | 班次 (Shift - 10月官方排班) |



            | :---: | :---: | :--- | :--- | :--- | :---: |



            | 1 | **ALA** | America Alanis | `aalanis` | Microbiologist I | **1st Shift** (5am-2pm) |



            | 2 | **ANC** | Andrew Carrillo | `acarrillo` | Microbiology Technician | **N/A (无班次)** |



            | 3 | **AOD** | Ayomide Odugbesi | `aodugbesi` | Microbiologist II | **2nd Shift** (9am-6pm) |



            | 4 | **CCD** | Cuong Du | `cdu` | Microbiologist I | **1st Shift** (5am-2pm) |



            | 5 | **CGS** | Clea S. Garza | `cgarza` | Microbiologist II | **1st Shift** (FMLA请假至11/11) |



            | 6 | **DT** | Debrework Tassew | `dtassew` | Microbiologist I | **1st Shift** (5am-2pm) |



            | 7 | **DK** | Dhvanir Kansara | `dkansara` | Project Microbiologist | **3rd Shift** (周末排班轮值) |



            | 8 | **ELB** | Ellen Bennett | `ebennett` | Microbiologist I | **2nd Shift** |



            | 9 | **EN** | Elysse Nioupin | `enioupin` | Microbiology Technician | **N/A (无班次)** |



            | 10 | **ES** | Alex Saravia | `asaravia` | Microbiologist I | **3rd Shift** (1pm-10pm) |



            | 11 | **GA** | Gerald Anyangwe | `ganyangwe` | Microbiologist I | **2nd Shift** (9am-6pm) |



            | 12 | **GL** | Guanchen Li | `gli` | Microbiologist II | **3rd Shift** (1pm-10pm, 严禁 David Li) |



            | 13 | **GS** | Gabrielle Surber | `gsurber` | Microbiologist II | **1st Shift** (5am-2pm) |



            | 14 | **HS** | Halaina Smith | `hsmith` | Microbiologist I | **1st Shift** (5am-2pm) |



            | 15 | **INS** | Inaara Sunesara | `isunesara` | Microbiologist I | **3rd Shift** (1pm-10pm) |



            | 16 | **ISS** | Ishita Sharma | `isharma` | Microbiology Technician | **N/A (无班次)** |



            | 17 | **JOC** | Jorge Ochoa | `jochoa` | Microbiologist I | **3rd Shift** (1pm-10pm) |



            | 18 | **KSM** | Karla Silva | `ksilva` | Microbiologist I | **2nd Shift** (7am-4pm, 严禁 Machuca) |



            | 19 | **MN** | Maryam Naeem | `mnaeem` | Microbiologist II | **1st Shift** (5am-2pm, 确认MN) |



            | 20 | **MJ** | Mukyung Jang | `mjang` | Microbiology Technician | **N/A (无班次, 严禁 Min Jang)** |



            | 21 | **MRB** | Muralidhar Bythatagari | `mbythatagari` | Microbiologist II | **3rd Shift** (1pm-10pm, 严禁 Reddy) |



            | 22 | **NH** | Neha Hulyalkar | `nhulyalkar` | Microbiology Supervisor | Supervisor |



            | 23 | **NT** | Nicholas Trinh | `ntrinh` | Microbiology Technician | **N/A (无班次, 确认NT)** |



            | 24 | **OA** | Olugbenga Ajayi | `oajayi` | Project Microbiologist | **1st Shift** (周末排班轮值) |



            | 25 | **PG** | Pagan Gary | `pgary` | Microbiologist II | **1st Shift** (5am-2pm) |



            | 26 | **QYC** | Qiyue Chen | `qchen` | Project Microbiologist | **1st Shift** (周末排班轮值) |



            | 27 | **RE** | Rey Estrada | `restrada` | Microbiologist I | **1st Shift** |



            | 28 | **RS** | Robin Seymour | `rseymour` | Microbiology Supervisor | Supervisor |



            | 29 | **SAA** | Sasha Allen | `sallen` | Microbiology Technician | **N/A (无班次, 确认SAA)** |



            | 30 | **SMO** | Simin Mohammad | `smohammad` | Microbiology Technician | **N/A (无班次)** |



            | 31 | **SU** | Sonal Uprety | `suprety` | Microbiologist II | **2nd Shift** (9am-6pm) |



            | 32 | **TK** | Tamiru Kotisso | `tkotisso` | Microbiologist I | **3rd Shift** (1pm-10pm) |



            | 33 | **VV** | Varsha Subramanian | `vsubramanian` | Microbiologist II | **2nd Shift** (9am-6pm) |







            - **历史在册/其他部门关联人员记录**:



              - `KA` $\to$ Kathleen Aruta (`karuta`) | `DH` $\to$ Domiasha Harrison (`dharrison`) | `DS` $\to$ Devanshi Shah (`dshah`)



              - `JE` $\to$ Jacqueline Esqueda (`jesqueda`) | `JM` $\to$ Jinal Mehta (`jmehta`) | `MC` $\to$ Maraya Chukwumerije (`mchukwumerije`)



              - `SAS` $\to$ Sophia Santamaria (`ssantamaria`) | `TD` $\to$ Thinh Duong (`tduong`) | `DPP` $\to$ Preeti Purwaha (`ppurwaha`)



              - `KSP` $\to$ Kathan Parikh (`kparikh`) | `RC` $\to$ Ross Caputo (`rcaputo`) | `JP` $\to$ Jay Patel (`jpatel`)



        *   **3. 底层代码工程强制防御与自动拦截机制 (Codebase Single Source of Truth & Interceptor):**



            - 在 `utils.py` 中建立全项目唯一的权威映射中枢（`get_full_name`, `clean_analyst_name`, `resolve_analyst_identity`）。



            - `pages/ScanRDI.py`、`pages/Celsis.py`、`pages/USP71.py` 及后台脚本必须统一从 `utils.py` 引用，严禁私自硬编码。



            - 在 `clean_analyst_name()` 与解析器中设置强制拦截钩子：任何包含 `eniola` 的字符串强制自动纠正为 `Elysse Nioupin`，任何 `gabbie` 强制纠正为 `Gabrielle Surber`，任何 `min jang` 强制纠正为 `Mukyung Jang`，任何 `elizabeth bennett` / `elizabeth` 强制纠正为 `Ellen Bennett`，从底层代码构筑无法击穿的技术防火墙。



        *   **4. 全文一致性校验与交付前门禁 (Universal Consistency & Pre-Delivery Audit Gate):**



            - 在 Page 1 分析员栏 (`Text Field3`)、访谈记录 (`Text Field13`)、正文各段落（Page 3/4/5）及独立台账表格中，分析员全名与缩写必须 100% 保持完全统一，交付前通过 `assert_no_hallucinations()` 自动审计，发现未授权名称直接阻断生成！



        *   **5. 🚨 绝对严禁自作聪明 · 不懂必问铁律 (Mandatory Ask-First Rule on Any Ambiguity / Unknowns - 终极行为准则):**



            - **零脑补 (Zero Guessing)**: 严禁自作聪明！无论是分析员名字、缩写、耗材批号、设备编号、培养箱校准日期、还是洁净室房间拓扑，遇到任何在权威字典和台账中查不到、模糊或不确定的信息，**绝对严禁自己猜或用大模型补全**！



            - **立即停步请示 (Stop & Ask Immediately)**: 只要遇到未知账号或缺失信息，必须**立即停步，直接向用户提问确认**。宁可多问一句，绝不乱写一字！



    47. **Robin Seymour (RS) 官方附表灰色分段横幅标准 (Table 2 Section Divider Golden Standard - 🚨 终极锁定):**



        *   **背景来源**: QA Reviewer Robin Seymour (RS) 在审阅并签署的官方标杆报告（如 `OOS-262194`）中，明确规范了 Table 2 中灰色分段横幅（Section Divider Band）的固定措辞，严禁随意拼接多余后缀或遗漏房间号。



        *   **分段横幅四大标准语法 (严格逐字执行)**:



            1. **人员监测分段**: 严格命名为 **`Personnel EM`**（🚨 **严禁追加日期，严禁写成 "Personnel EM for 16Sep26"！**）



            2. **生物安全柜分段**: 严格命名为 **`Biological Safety Cabinet EM for E00XXXX and E00YYYY`**（🚨 **严禁出现 "(BSC)"，严禁重复出现 "BSC"，严禁追加 "for [Date]"！** 例如：`Biological Safety Cabinet EM for E001938 and E001937`）



            3. **洁净室浮游菌分段**: 严格罗列该套间全部房间号与设备编号，格式为 **`Weekly Active Air Sampling of Cleanroom [房间号清单] ([设备编号], [套间名称]) with [Processing/Changeover] BSCs`**（🚨 **严禁追加 "for [Date]"！** 例如：`Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (E001979, L-Suite) with Processing BSCs` 或 `Weekly Active Air Sampling of Cleanroom Suite 116 (E001738) with Processing BSC`）



            4. **洁净室表面分段**: 严格与浮游菌保持相同的套间房间号标准，命名为 **`Weekly Surface Sampling of Cleanroom [房间号清单] ([设备编号], [套间名称]) with [Processing/Changeover] BSCs`**。



    48. **Phase I 结尾样品历史回溯与开启 Phase II 过渡段标准 (Phase I Sample History Discovery & Phase II Initiation Closing Standard - 🚨 终极锁定):**



        *   **核心逻辑**: 当某个 OOS 案件包含 Phase I 和 Phase II 两份调查报告（例如初始 USP <71> 为阳性，但调查批号历史时发现同一批次此前提交的 ScanRDI 为阴性合格），**不仅要在 Phase II 中交代背景，更必须在 Phase I 第 5 页 (`Text Field51`) 的最末尾加上过渡收尾段，并且必须明确写出复检样品的 ETX 编号（如 `ETX-260914-0470`）**！



        *   **Phase I 末尾标准句式**:



            `"However, when investigating the lot history, it was discovered that the same lot (Lot [LOT_NUMBER]) had previously been submitted for ScanRDI sterility testing under [SCAN_ETX_ID] and yielded a passing result. Based on this conflicting passing ScanRDI result for the same lot, the laboratory requested additional sample vials from the client for retesting under [RETEST_ETX_ID] and initiated a Phase II investigation (Form 3.100.019.F02)."`



        *   **前后衔接一致性与紧凑排版防留白规范**:



            - Phase I 的最末尾与 Phase II 的开头必须形成 100% 因果闭环（系实验室内部在调查批号历史时主动发现同批次合格记录并主动向客户请求复检样品 `under [RETEST_ETX_ID]`，严禁编造“客户主动指出”或遗漏复检 ETX 号）。



            - **Page 3 / 4 / 5 紧凑排版防留白与防截断法则**: Page 3 (`Text Field49`) 排入 `p1..p8` 共 8 段（字号校准为 `8.65pt` 以确保第 8 段末尾 `"to begin incubation."` 100% 完整显示绝不截断），Page 4 (`Text Field50`) 排入 `p9..p18` 共 10 段（`9.2pt`），Page 5 (`Text Field51`) 排入 `p19..p21` 共 3 段（`9.2pt`），彻底消除页面底部多余留白。







    49. **Standard Microbiology SOP Designation Rule (官方规程编号规范 - MICRO-SOP-2 - 🚨 终极锁定):**



        *   **标准编号**: 洁净室环境监测培养与设施监控规程必须统一规范命名为 **`MICRO-SOP-2`**（`Environmental Monitoring of the Clean-room Facility`）。



        *   **严禁旧编号**: 在所有 OOS 调查报告（包括 ScanRDI、Celsis、USP <71>、EM）的正文叙述及表格说明中，**绝对严禁再写成旧编号 `SOP 2.600.002`**！必须统一写作 `MICRO-SOP-2`！







    50. **Robin Seymour (RS) Benchmark Golden Standard Harmonization for OOS Reports (Robin 质保复核黄金标杆与全 OOS 模块通用八项铁律 - 🚨 终极锁定):**



        *   **标杆基准与背景**: 以 QA Director Robin Seymour (RS) 官方复核并签署归档的最终版 OOS 调查报告（标杆如 `OOS-262150 BelieveRX Pharmacy Signed by QYC` 及 `OOS-262194`）为绝对依据，提炼并固化全 OOS 模块（ScanRDI, Celsis, USP <71>, EM）通用的高质量排版与行文规范：



        *   **1. Page 4 处理与换台同人员合并句式 (Combined Processing & Changeover Phrasing):**



            - 当同一位分析员同时执行了样品过滤处理（Processing）与换台步骤（Changeover）时，正文叙述严禁拆分成两句机械重复的话（如 `"processed by X, and changeover performed by X"`）。



            - 必须使用 Robin 标准合并从句：`"processing and change over steps were performed by [Analyst] ([Nth] sample processed)"`（例如：`"processing and change over steps were performed by Ellen Bennett (4th sample processed)"`）。



        *   **2. Page 3 容器完整性段落分析员连字符规范 (Analyst Name Hyphen Convention):**



            - 在第 2 段（容器完整性与样品外观检查）中，分析员全名两侧必须添加空格连字符包裹：



              `"Prepping analyst - [Name] - and processing analyst - [Name] - verified the physical integrity and packaging of the sample..."`



              （例如：`"Prepping analyst - Mukyung Jang - and processing analyst - Ellen Bennett - verified..."`），严禁使用无连字符的直白裸名。



        *   **3. Page 4 环境监测评估三大区块排版层次与原生回车标准 (EM Recoveries Review Block Hierarchy):**



            - 在 AcroForm 多行文本框（`Text Field50`）中，环境监测评估部分必须清晰拆分为三大由 `\r \r` 严格隔开的独立区块，且小标题与正文之间使用单 `\r` 紧凑连接：



              - 区块 1 (主标题独立成段): `"Environmental Monitoring (EM) Recoveries Review:"`



              - 区块 2 (BSC 与人员评估): `"ISO 5 BSC and Personnel Evaluation:\r[body text]"`



              - 区块 3 (洁净室评估): `"Cleanroom EM Evaluation:\r[body text]"`



            - 小标题与正文之间使用单个 `\r` 换行，能确保小标题与段落紧密相随，同时杜绝在 Adobe Acrobat 中产生双重空行或行距漂移。



        *   **4. Page 2 设备栏全量罗列生物安全柜及其校准期 (Page 2 Equipment Block Must Include BSCs):**



            - 在 Page 2 `Text Field43`（Equipment ID）与 `Text Field44`（Calibration Due Date）中，除了培养箱与温度传感器之外，**必须完整列出处理所用的生物安全柜及换台生物安全柜**（例如：`BSC E001938` 与 `BSC E001937`），并在右侧严格对应其校准到期日（如 `Feb 2027` 与 `Feb 2027`）。



            - 每一组设备块之间必须保持严格的 `\r \r` 空行隔开与横向 1:1 绝对对齐。



        *   **5. Page 2 耗材批号/效期原生回车规范 (Text Field22 Reagent Formatting):**



            - 耗材与试剂批号/效期栏目必须采用 Acrobat 原生回车换行符（同一组内使用 `\r`，培养基与耗材大类之间使用 `\r \r`），严禁混入裸 `\n`，杜绝因换行符解析不一致导致 Adobe Acrobat 行高塌缩或末尾截断。



        *   **6. Page 6 调查撰写人栏目纯粹性 (Text Field53 Prepared By Print Standard):**



            - `Text Field53`（Prepared by Print）必须纯粹填写调查报告撰写人/调查员姓名（如 `"Qiyue Chen"`），严禁与实验操作人员混搭拼接（严禁写成 `"Ellen Bennett (Written by: Qiyue Chen)"`）。



        *   **7. Table 2 灰色分段横幅标准与日期后缀调和 (Section Divider Banners Standard):**



            - 官方终极基准：



              1. 人员监测: `Personnel EM`（无日期后缀为官方主基准；在单日追溯特定案卷中，若 QA 明确要求带日期，则精确呈现为 `Personnel EM for DDMMMYY`）。



              2. 生物安全柜: `Biological Safety Cabinet EM for E00XXXX and E00YYYY`（或 `(BSC) Biological Safety Cabinet EM for BSC E00XXXX and BSC E00YYYY`）。



              3. 洁净室浮游菌: `Weekly Active Air Sampling of Cleanroom [Rooms] ([Equipment], [Suite]) with Processing BSCs`（无日期后缀）。



              4. 洁净室表面: `Weekly Surface Sampling of Cleanroom [Rooms] ([Equipment], [Suite]) with Processing BSCs`（无日期后缀）。



        *   **8. Table 2 周检周期列纯净表头 (Day/Week Column Header Standard):**



            - Table 2 周检行的 `Day /Week(s)` 列必须严格填写为 `"Week of Testing"`，严禁追加 `"\n(Scan)"` 或 `"Week on Testing Date"`。



        *   **9. 客户 6 个月历史阳性次数严格真实审计 (Confirmed Prior OOS Numerator Audit):**



            - 在核查客户 6 个月历史记录时，分子（先前 OOS 发生次数）仅能计入真实确认有效的无菌检验失败事件。对于被撤回（Canceled）、作废或判为无效的历史 OOS 事件，严禁计入分子虚增客户阳性率（例如 BelieveRX E74120 确证历史阳性事件为 3 起，而非 5 起）。







    51. **Scan RDI OOS-262237 (GoGoMeds Select E10747) 现场数据更新与 Robin Seymour 标杆完全对齐 (Scan RDI OOS-262237 Finalization & Benchmark Alignment - 🚨 终极锁定):**



        *   **案卷基本信息**:



            - OOS 编号: `OOS-262237`



            - 客户: `GoGoMeds Select (E10747)`



            - 样品编号: `ETX-260921-0148` (Semaglutide R 2.5mg/ml Benzyl Alcohol 0.9%, Lot: 20066, Injectable)



            - 检验日期: `24-Sep-2026`，检验记录: `092426-2132-3`



            - 结果: 362 total events, 90 confirmed microbial events (curved rod-shaped morphology)



            - 核心人员: Prepper = Elysse Nioupin (EN), Processor = Guanchen Li (GL, 第 27 支样品), Changeover & Reader = Muralidhar Bythatagari (MRB)



            - 设备: Scan RDI E002132 (Oct 2026), BSC E001312 (Jun 2027, CR116B), BSC E001937 (Jul 2027, CR144)



            - 阳性对照法定台账: `C. albicans`, Lot: `06122026-10231-CA`, Exp: `11Jun28` (GL 30, Session: 092426-2225-3)



        *   **EagleTrax 现场 EM 数据自主穿透查询与确证**:



            - 浮游菌 `ETX-261005-0773` (Suite 116 Anteroom, 3 CFUs): 3 菌落均为 **`Gram (+) cocci`**



            - 浮游菌 `ETX-260929-0335` (L-Suite ISO 8 143 Sec I, 6 CFUs): 5 菌落 **`Gram (+) cocci`**，1 菌落 **`Gram (-) coccobacilli`**



            - 浮游菌 `ETX-260929-0341` (L-Suite ISO 8 143 Sec II, 8 CFUs): 6 菌落 **`Gram (+) cocci`**，2 菌落 **`Gram (+) short rods`**



            - 浮游菌 `ETX-260929-0344` (L-Suite ISO 8 142, 2 CFUs): 2 菌落均为 **`Gram (+) cocci`**



            - 关键防守逻辑：背景 ISO 8 检出的球菌、球杆菌与短杆菌形态与待测品中检出的 **弯曲杆菌 (curved rod-shaped morphology)** 具有鲜明形态学差异，且关键 ISO 5 工作面及操作员手套无菌生长（absence of microbial recovery），彻底切断外围污染转移路径。



        *   **Robin Seymour (RS) 标杆全要素对齐 (`OOS-262150 BelieveRX Signed by QYC` 标准)**:



            - **Page 1 字段精细化**: `Text Field0` 格式化为 `Muralidhar Bythatagari (Written by: Qiyue Chen)`（字号 5.12pt），Section B 问卷评语全量填入并应用基准字号。



            - **Page 2 培养箱、设备与洁净室规范**: 



              - `Text Field43` 培养箱与传感器分行（`Incubator E001034\r \r(Sensor E001501)\r \r...`），`Text Field44` 校准期分行（`Aug 2027\r \rFeb 2027\r \r...`），右侧追加 BSC E001312 (`Jun 2027`) 与 BSC E001937 (`Jul 2027`)，字号严格设为 `8.5pt`，实现 1:1 绝对横向基准线对齐。



              - `Text Field22` 与 `Text Field23` 耗材栏字号设为 `6.0pt`。



              - `Text Field24`、`25`、`26` 严格填入官方确证阳性对照：`C. albicans`、`06122026-10231-CA`、`11Jun28`。



              - `Text Field32`、`33` 洁净室认证多套间分行规范：`E001738 (CR116)\rE001978 (CR144)` 与 `Dec 2026\rDec 2026`（字号 4.4pt，具体到房间代号，**严禁使用 "&" 连接符**）。



            - **Page 3 / 4 / 5 正文紧凑版面流转**:



              - Page 3 (`Text Field49`, 9.2pt): 样品前处理、双重漂白水消毒、套间压差、BSC 认证、过滤前设置、检验结果描述及 Table 2 引入（p1 至 p9，规范使用 `MICRO-SOP-2`）。



              - Page 4 (`Text Field50`, 8.1pt): 严格以 `Environmental Monitoring (EM) Recoveries Review:` 起首（p10 至 p17）：



                - ISO 5 与人员评估：严格去除 "100%" 绝对化用语，标准表述为 `absence of microbial recovery (No Growth) was demonstrated across all critical ISO 5 zones and personnel monitoring on the date of testing`。



                - 洁净室环境评估：周检浮游菌语法修正为 `in the ISO 7 cleanrooms (145 and 144)`，外围 ISO 8 检出与待测样形态对照排查及压差屏障防守；ISO 8 暴露技术防守用语统一规范为 `"There is minimal exposure of the test sample to ambient ISO 8 air, as samples and supplies are transferred in disinfected, closed containers on carts through the layered cleanroom suites."`（绝不可写成绝对化的 "does not come into contact with ambient ISO 8 air"）。



                - 月度清洁：明确核实并精确拆分记录为 `16Aug26 for CR116 and 30Aug26 for L-Suite`，所有化学指示卡合格。



                - 6 个月客户批次历史真实审计与三步曲闭环：确证存在 1 起先前失败事件 `OOS-262106`（检验日期 10Sep26，样品 ETX-260908-0694，批号 19622，GL/MRB）。严格执行 Robin Seymour 审核通过的三步曲闭环规范：



                  1. 基础案卷陈述句（加逗号隔开）：`Analyzing a 6-month sample history for GoGoMeds Select (E10747), this specific analyte "Semaglutide R 2.5mg/ml Benzyl Alcohol 0.9%" has had one prior failure, documented under OOS-262106 for submission ETX-260908-0694, tested on 10Sep26.`



                  2. 形态学差异对比句：`The microorganism detected in the previous investigation exhibited circular-shaped (clusters) morphology, which differs from the curved rod-shaped morphology observed in the current test sample.`



                  3. 前案回顾闭环评价句：`The previous failure investigation was reviewed and no laboratory error or assignable laboratory cause was identified.`



                - 交叉污染排查与结论：



                  - 交叉污染开篇句统一规范为：`"All samples processed on the same day were also evaluated for evidence of sample-to-sample cross-contamination."`



                  - 第 27 支样品，全天其余样品全阴性，最终结论法定句统一为：`"Therefore, the likelihood of laboratory error contributing to this failure is considered minimal, and the original result is deemed valid."`



              - Page 5 (`Text Field51`, 11.0pt): 严格遵循标杆仅填写 **`N/A QYC 13Oct2026`**，避免跨页断行与底部多余留白。



            - **Page 6 签字栏**: Section D 复选框选择 `Check Box88`（'No, the OOS result is valid.'），调查人 `Text Field53` 规范填写 `Qiyue Chen`（7.5pt）。



            - **Page 7 附表集成**: Standalone Table 1 & Table 2 完美整合为第 7 页，内嵌 Sample ID 与全部阳性 EM 培养皿的 live 原生可点击超链接，总页数严格锁定为 7 页。







    52. **Native Auto Font Size (fs=0) for ALL PDF Form Transcriptions (全 OOS 模块 PDF 表单转录原生 Auto 字号永久铁律 - 🚨 终极锁定):**



        *   **🚨 血的教训与根本病灶剖析 (Root Cause of "Locked Font Size" Bug):** 在 OOS-262098 Phase II 转录过程中发现：旧版转录脚本使用 PyMuPDF (`fitz`) 遍历所有文本控件并强制设置 `w.text_fontsize = <固定值>`（如 8.5, 9.2, 5.8, 3.8 pt），这会将 PDF 原生 `/DA` 字符串从 `Helv 0.000 Tf`（Auto 自适应）覆盖为 `Helv 8.5 Tf`（固定死字号）。一旦字号被锁死，用户在 Foxit/Acrobat 中编辑内容（如添加空行或修改文字）时，文本将无法动态缩放，丧失了人类手动在 PDF 中复制粘贴的原生体验。



        *   **1. 绝对禁止硬编码固定字号 (Strict Ban on Hardcoded Font Sizes via fitz):**



            - 在所有 OOS 模块（`ScanRDI`、`Celsis`、`USP <71>`、`EM`）的 PDF 表单转录脚本中，**严禁通过 `fitz` / PyMuPDF 的 `w.text_fontsize = <任何非零数值>` 来设置文本框字号！**



            - 仅允许对 Checkbox 控件执行 `fitz` widget 操作（勾选/取消勾选），所有文本字段的值写入必须且仅能通过 `pypdf` 的 `update_page_form_field_values()` 完成。



        *   **2. 强制启用 NeedAppearances (Mandatory NeedAppearances Flag):**



            - 在 pypdf `PdfWriter` 中，必须调用 `writer.set_need_appearances_writer(True)`（或等效方法），指示 PDF 阅读器（Foxit/Acrobat）在打开文档时自动重新生成外观流（Appearance Stream），从而实现字体大小动态自适应。



        *   **3. 如必须操作 fitz 文本控件则强制设为 0 (If fitz Text Widgets Must Be Touched, Force fs=0):**



            - 在极端特殊场景下（如无法完全避免 fitz 文本控件操作），**必须显式设置 `w.text_fontsize = 0`（Auto）**，绝不允许设置为任何正数固定值。



        *   **4. 官方空白模板验证基准 (Official Blank Template Baseline):**



            - 经验证，ZenQMS 官方下载的空白表单（`CORP-FORM-21` 与 `CORP-FORM-22`）所有文本字段原生均为 `fs=0.0`（Auto）。转录脚本的目标是 100% 保留这一原生状态，使最终产出与人类在 Foxit 中手动复制粘贴的效果完全一致。



        *   **5. 全模块全阶段覆盖 (Universal Scope — All Modules, All Phases):**



            - 此规则适用于所有当前与未来的 OOS 调查报告转录流程，包括但不限于：



              - Phase I 转录（`CORP-FORM-21 / 3.100.019.F01`）



              - Phase II 转录（`CORP-FORM-22 / 3.100.019.F02`）



              - EM 转录、独立表格 PDF 生成等



    53. **Phase II Retest Investigation & Exact Phase I Verbatim Transcription Standard (Phase II 调查报告与 Phase 1 叙述逐字逐句 100% 镜像转录铁律 - 🚨 终极锁定):**



        *   **🚨 业务认知与病灶剖析 (The Verbatim Transcription Mandate):** 在起草或转录 Phase II 调查报告（`CORP-FORM-22 / 3.100.019.F02`）时，第 2 页至第 4 页的叙述核心为两阶段对比调查。在 `INITIAL TEST UNDER SUBMISSION [ETX-XXXXXX-XXXX]:` 标题下的内容，**必须 100% 原汁原味完整复制转录已结案定稿的 Phase I 调查叙述**（包括其完整段落结构、换行符、标点符号、空行等），严禁自行二次总结、改写、缩减或遗漏任何一段！



        *   **两阶段严格分界准则 (Strict Section Demarcation):**



            - 初始检验部分：严格收录在 `INITIAL TEST UNDER SUBMISSION [ETX-XXXXXX-XXXX]:` 之后，完整呈现 Phase I 调查事实。



            - 复测检验部分：必须严格以 `PHASE II RETEST UNDER SUBMISSION [ETX-XXXXXX-XXXX]:` 作为独立标题另起炉灶，仅在这一部分书写 Retest 的新流程、新人员、新设备、新环境与新结果。



        *   **非纯连续正文空行美学规范 (Empty Line Visual Ergonomics):** 在标题与段落、不同检验阶段、或列表式结构块之间，必须空出一行（使用 `\r\n\r\n` 或两个换行符），在 PDF 自适应字号（Auto Font Size）下形成优美疏朗的视觉层级，绝不可密密麻麻挤作一团。







    54. **Elimination of Robotic AI All-Caps Headers & Over-Explaining Conciseness Rule (严禁机械全大写 AI 标题与“写得越多越错”精简防御铁律 - 🚨 终极锁定):**



        *   **严禁大写 AI 式标题 (Strict Ban on Robotic All-Caps Headers):** 调查报告是供 QA 和监管机构审计的法定技术文档，**严禁使用诸如 `EVALUATION OF RETEST FINDINGS AND JUSTIFICATION OF PRIOR SCANRDI RESULT:` 等生硬刻板的全大写 AI 风格小标题**。段落承接必须像经验丰富的资深分析员一样自然流动，使用自然的英文过渡句（如 `In evaluating the prior rapid ScanRDI testing for Lot ...`）。



        *   **“写得越多越错”精简防御铁律 (Conciseness Over Speculative Essays):**



            - 当解释为什么先前同批次快速法（如 ScanRDI）合格而后续药典法（USP <71>）复检不合格时，**严禁脑补写成大篇幅的学术辩解论文**（例如大谈特谈取样体积差异、没有方法学验证、单瓶与多瓶抽检的概率论二项分布、肉汤增菌动力学等）。在 cGMP 审计中，写得越多，给审查官暴露的抓柄和合规漏洞就越多！



            - **法定精简防御用语**: 仅需给出科学严谨且无法被反驳的客观核心定论即可：



              `"...attributed to non-uniform microbial distribution at low concentrations within the compounded batch."`（归因于低浓度微生物在复配批次中的非均匀分布）。一句话直击要害，绝不多写一句臆测之词！







    55. **Preservation of Native Digital Signature Fields in PDF Forms (CORP-FORM-21 / CORP-FORM-22 原生数字签名域绝对禁止转为普通文本铁律 - 🚨 终极锁定):**



        *   **🚨 严重失误与底层病灶 (The Digital Signature Stripping Bug):** 在 `CORP-FORM-22` 第 5 页 `Section F – Acknowledgement and Closure`（以及所有 ZenQMS 官方 PDF 表单）中，虽然表头打印文字为 `Print Name:`，但底层的 PDF AcroForm 字段实际上是 **`/FT /Sig` 原生电子数字签名域（如 `Signature_0`, `Signature_1`, `Signature_2`）**。



        *   **破坏后果**: 若脚本为了强行在框中打印名字字符串（如 `Qiyue Chen`）而把底层控件类型修改为 `/FT /Tx`（Text 文本框），会导致 Foxit / Acrobat 的官方数字证书签名机制（紫框与红/橙色 "SIGN" 旗标）被彻底摧毁，用户将无法调用个人数字证书（Digital Certificate）进行具有法律效力的电子签章！



        *   **铁律规定 (Ironclad Mandate):**



            1. **严禁改写签名域类型**: 绝对严禁将任何 `/FT /Sig` 字段转换为 `/FT /Tx`，严禁向签名控件中强制填入纯文本字符串！



            2. **保持空白留待验签**: `Signature_0`（Prepared by:）必须保留为纯粹的原生空白 Signature 控件，以便调查员在 Foxit 中点击时自动调出数字证书盖章（生成 "Digitally signed by Qiyue Chen" 及时间戳元数据）。



            3. **仅填充对应日期**: 仅将调查当天日期填入对应的日期框（如 `Date Field3` 填入 `16-Oct-2026`）。



            4. **评审人与质保栏纯净留空**: `Reviewed by:`（`Signature_1`, `Date Field4`）与 `Approved by (QA):`（`Signature_2`, `Date Field5`）必须保持 100% 完全空白，等待后续主管与质保人员签署。







    56. **Phase II Section E Scope & Permission Protocol (Phase II Section E 调查范围与权限门禁规范 - 🚨 终极锁定):**



        *   **调查员权限边界**: 在 `CORP-FORM-22` 中，`Section E – Corrective Actions and Preventive Actions / Additional Evaluation` 通常属于质保（QA）或管理层在完成综合调查后的纠正预防措施制定权限范围。



        *   **默认留空防越权**: 除非质保主管或用户有极其明确的书面指令，调查员在起草报告时**绝不可擅自填写 Section E**。默认必须保持 Section E 完全留空，防止因擅自设定不当 CAPA 而触犯合规红线。







    57. **Proactive Official ZenQMS Form Transcription Reminder & Workflow (主动提醒/一键触发官方表单转录常态化铁律 - 🚨 终极锁定):**



        *   **业务痛点与用户诉求**: 每一起 OOS 调查案件立案时，质保（QA/ZenQMS）都会下发一份全新的当日有效空白官方 PDF 表单（Phase I 为 `CORP-FORM-21 / 3.100.019.F01`，Phase II 为 `CORP-FORM-22 / 3.100.019.F02`）。手动在 PDF 中逐格填入 150+ 个表单字段和勾选数十个复选框极其繁琐耗时，且极易遗漏。



        *   **强制主动提醒门禁 (Mandatory Proactive Reminder Gate):**



            - 在任何 OOS 调查案件的检测数据梳理、叙述初稿起草、或 Word 报告生成完毕后，**AI 必须在回复的最末尾主动提醒用户进行官方表单转录**！



            - **标准主动提醒话术**:



              `"💡【官方表单转录提醒】：检测叙述与台账数据已全部核验完成！请问你今天从 ZenQMS 下载了最新空白官方 PDF 表单（CORP-FORM-21 或 CORP-FORM-22）吗？如果有，请直接将文件或路径告诉我，我将立即为你自动执行 100% 官方表单转录、自动字号匹配（Auto Font Size）与附件表格合成，省去手动复制粘贴的繁琐操作！"`



            - 若用户已将空白表单放置于桌面或工作目录，AI 应当主动检测并提议直接执行一键转录。







    58. **Official EM Active Air OOS Investigation Golden Benchmark (OOS-261553 Standards Harmonization - 🚨 终极锁定):**



        *   **标杆基准与背景**: 以质保（QA）及实验室主管 Neha Hulyalkar 官方复核并签署归档的最新版本环境监测（EM Active Air）OOS 终极法定标杆案卷 `OOS-261553 EM SMO 142 Air 26JUN2026.pdf`（CORP-FORM-21 v11.1）为绝对标准，固化全流程 EM OOS 案卷规范：



        *   **1. Page 1 Initiator 官方代笔命名规范 (Initiator Authorship Convention):**



            - 当由调查员/撰写人为检测分析员代笔起草调查报告时，`Text Field0`（Initiator）必须严格采用官方批准的标准代笔格式：`"[Investigated Analyst] (Written by: [Author])"`（例如：`"Simin Mohammad (Written by: Maryam Naeem)"` 或 `"Simin Mohammad (Written by: Qiyue Chen)"`），杜绝责任人与起草人混淆。



        *   **2. Page 1 事件极简客观陈述 (Standard Description of Incident):**



            - `Text Field7`（Description of Incident）必须采用经质保审核的标准精简客观句：`"The CFU count for the environmental monitoring plate exceeded the action level."`，严禁加入未经调查的推测性言论。



        *   **3. Page 1 SOP 编号与限度值规范 (SOP & Specification Standard):**



            - SOP 引用严格填写最新生效版 `MICRO-SOP-2`（`Rev. 16`，Effective `23 Jul 2026`）。



            - 限度栏目（`Text Field11`）严格区分 Alert 与 Action Level：



              `Alert Level: ≥ 8 CFU/Plate\rAction Level: > 10 CFU/Plate`



        *   **4. Page 2 环境监测符合性严格勾选 "No" 铁律 (EM Conformity Checkbox Mandatory "No"):**



            - 针对 `Do the environmental monitoring results conform to required standards?`，必须严格勾选 **No**（即在表单 No 列的 `Check Box68` 勾选 `Yes`），并在对应的数据栏填入 `No` / `No`。因为本案调查对象本身即为环境监测超标事件，绝不可自相矛盾勾选“符合标准”。



        *   **5. Page 2 培养箱校准到期日锁定 (Incubator Aug 2027 Invariance):**



            - 培养箱 E001031 与 E001034 在 2026 年度的校准到期日严格为 **`Aug 2027`**，内置温湿度传感器 E001505 与 E001501 严格为 **`Feb 2027`**。在 `Text Field43` 与 `Text Field44` 中必须严格保持 1:1 水平空行基准线对齐。



        *   **6. 全表单 AcroForm 原生 Auto 字号 (fs=0.0) 与人类排版体验 (PRD Rule 52 深度实践):**



            - 全表单所有文本框必须保持 `fs=0.0`（Auto Font Size），正文段落之间使用原生 `\r \r` 分隔。



            - Page 3（`Text Field49`，6 段）、Page 4（`Text Field50`，7 段）、Page 5（`Text Field51`，3 段）自然流动排布，消除多余机械留白与截断，体验与 Foxit PDF Editor 手工录入完全一致。



        *   **7. Page 4 洁净室套间月度清洁精细化房间罗列规范 (Multi-Room Cleaning Specificity):**



            - 月度清洁必须逐一明确列出各级房间编号与等级：如 `ISO 8 Anteroom (CR142), ISO 8 Anteroom (CR143), ISO 7 Buffer Room (CR144), ISO 7 Buffer Room (CR145), and associated ISO 5 Biosafety Cabinets`，明确清洁日期（如 `14 Jun 2026` 或 `30 Aug 2026`）、操作人及双氧水指示卡合格情况。



            - 必须明确包含产品排查陈述：`"Notably no samples processed within that week in [Room] failed ScanRDI test."`



        *   **8. Page 7 独立附表单页上下集成标准 (Single-Page Integrated Table 1 & 2):**



            - 案卷第 7 页必须单页整合 `Table 1: Read Dates & Incubation Observation`（8 列）与 `Table 2: Environmental Monitoring for Analyst & Cleanroom`（9 列），周期列严格使用 `Week (Of Testing Date)` 与 `Week (After Testing Date)`，杜绝无意义跨页。



        *   **9. 独立立案与附表绝对解耦原则 (Strict Single-OOS Isolation & Table 1 Disaggregation Rule):**



            - 凡在同一采样日、同一套间中出现多洁净室或多分区同时超标并分别获得独立 OOS 案卷编号时（如 OOS-262133 对应 CR143，OOS-262134 对应 CR142）：



                - **各自建立独立案卷**：每个 OOS 案卷必须拥有其专有的 PDF 报告包与专有的单页 Word/PDF 附表（如 EM table OOS-262133 04SEP2026.docx 与 EM table OOS-262134 04SEP2026.docx）。



                - **Table 1 样品绝对归属**：Table 1 (Read Dates & Incubation Observation) **严禁将其他 OOS 案卷的样品混合填入**！Table 1 只能且必须只列出本 OOS 案卷管辖范围内的 ETX ID 与 CFU 读数。



                - **Table 2 环境回溯保留全局参考**：Table 2 (Environmental Monitoring for Analyst & Cleanroom) 仍应保留同套间其他房间的监测记录以提供完整的环境背景与趋势回溯，确保 bracketing 的完整性。



                - **PDF 附表页严格对应**：每个 7 页 PDF 案卷的 Page 7 只能挂载属于自己案卷的独立附表，杜绝跨案卷串门。







59. **Scan RDI OOS-262289 Master Production Package & Same-Day ZenQMS Form Transcription (Link Pharmacy E19112 - 铁律闭环):**

    *   **案件背景 (Background):**
        - 客户与样品: Link Pharmacy (E19112), Sample `ETX-260929-0931`, `Tirzepatide/Levocarnitine 17mg/100mg/mL (1ml)`, Lot `20260928@54`, Dosage Form: Injectable.
        - 测试方法与设备: Scan RDI Sterility Test (`MICRO-SOP-12`, `ENG-SOP-4`), Instrument `E001040` (Cal Due `Oct 2026`), L-Suite Cleanroom 145 (`E001979`, Cert Due `Nov 2026`, Processing BSC `E001938` Cal Due `Jul 2027`) & Cleanroom 144 (`E001978`, Cert Due `Nov 2026`, Changeover BSC `E001937` Cal Due `Jul 2027`).
        - 结果 (RS Reviewed): **19 total events, 4 confirmed microbial events** exhibiting `Cocci shaped morphology` (verified per Scan Data packet).
        - 阳性对照与实验批次: `P. aeruginosa`, Lot `05162026-9027-PA`, Exp `15May28` (Session 100126-1040-2001, CONTROL + ELB).

    *   **终极法定附表数据 (Finalized Ground Truth Tables):**

        *   **Table 1: Information for ETX-260929-0931 under investigation**
            | Processing Analyst | Reading Analyst | Sample ID | Events | Confirmed Microbial Events | Morphology Description |
            | :--- | :--- | :--- | :--- | :--- | :--- |
            | Ellen Bennett | Varsha Subramanian | [ETX-260929-0931](https://etrax.eagleanalytical.com/Submission/Details/sM81r08$r4X9k5LrerlJ7g__) | 19 | 4 | Cocci shaped morphology |

        *   **Table 2: Environmental Monitoring from Testing Performed on 01Oct26 (全表数据单元格统一 7.0pt)**
            | Environmental Monitoring (EM) Sampling Site | Frequency | Date (DDMMMYY) | Analyst | Day /Week(s) | Observation | EM Plate ETX Number | Microbial ID | Notes |
            | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
            | **Personnel EM** | | | | | | | | |
            | Personal (Left Touch and Right Touch) | Daily | 01Oct26 | ELB | Date of Testing (Scan) | No Growth | Not Applicable | Not Applicable | None |
            | Personal (Left Touch and Right Touch) | Daily | 01Oct26 | KSM | Date of Testing (Scan C/O) | No Growth | Not Applicable | Not Applicable | None |
            | **Biological Safety Cabinet EM for E001938 and E001937** | | | | | | | | |
            | Surface Sampling of ISO 5 BSC E001938 (4 Locations) | Daily | 01Oct26 | ELB | Date of Testing (Scan) | No Growth | Not Applicable | Not Applicable | None |
            | Surface Sampling of ISO 5 BSC E001937 (4 Locations) | Daily | 01Oct26 | KSM | Date of Testing (Scan C/O) | No Growth | Not Applicable | Not Applicable | None |
            | Settling Sampling of ISO 5 BSC E001938 | Daily | 01Oct26 | ELB | Date of Testing (Scan) | No Growth | Not Applicable | Not Applicable | None |
            | Settling Sampling of ISO 5 BSC E001937 | Daily | 01Oct26 | KSM | Date of Testing (Scan C/O) | No Growth | Not Applicable | Not Applicable | None |
            | **Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs** | | | | | | | | |
            | Active Air Sampling of Cleanrooms | Weekly | 01Oct26 | SMO | Week of Testing | 1 CFU<br>(ISO 7 144 Sec II) | [ETX-261012-0430](https://etrax.eagleanalytical.com/Submission/Details/xxCNoUxIy4doEwKyIX-rYg__) | *Staphylococcus epidermidis* | None |
            | Active Air Sampling of Cleanrooms | Weekly | 01Oct26 | SMO | Week of Testing | 5 CFUs<br>(ISO 8 142) | [ETX-261012-0433](https://etrax.eagleanalytical.com/Submission/Details/Qr5FQ8nlCFY1Y9QOAyCjZQ__) | *Staphylococcus epidermidis*<br><br>*Pseudomonas parafulva*<br><br>*Bacillus safensis*<br><br>*Corynebacterium xerosis*<br><br>*Corynebacterium ureicelerivorans* | None |
            | **Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs** | | | | | | | | |
            | Surface Sampling of Cleanrooms | Weekly | 01Oct26 | SMO | Week of Testing | No Growth | Not Applicable | Not Applicable | None |

        *   **同日其他阳性批次评估对照数据 (Concurrent Positive Sample Evaluation):**
            | OOS ID | Client Name | Sample ID | Analyte Name | Lot # | Prepper | Processor | Suite & BSC | Morphology |
            | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
            | OOS-262285 | CasaPharma Rx (E72585) | ETX-260930-0635 | Tirzepatide/B12 15mg/1mg/0.5mL - 2mL | 260929E | Pagan Gary (PG) | America Alanis (ALA, 11th) | Parallel Suite (BSC 1312, C/O 1937) | cocci (cluster/tetrad) shaped morphology |

    *   **核心质保与合规执行准则 (Compliance & Golden Execution):**
        1. **同日下载 ZenQMS 官方表单转录铁律 (Same-Day ZenQMS Form Invariant):**
           - 严格转录至阳性发现当日下载之官方空白表单 `CORP-FORM-21- 01 Oct 2026.pdf`，确保文件头部 ZenQMS 加密审计水印 (`Printed by rseymour@eagleanalytical.com ... on 01-Oct-2026 at 11:12:23 PM UTC`) 真实合规不可篡改。
        2. **Section B 行位精准 1:1 对齐纠偏 (Section B 1:1 Alignment Fix):**
           - 彻底修复表单由于 `Text Field17` 错位导致的第 5 至 9 行偏移问题：`Text Field17` 准确填充 `"Yes, See 100126-1040-2 for more information."`，使 rows 5-9 各项检查说明严格对齐。
        3. **发起人字号精准微调 (Initiator Font Size Calibration):**
           - `Text Field0` (Initiator) 设为 `4.75pt`，确保 `"Varsha Subramanian (Written by: Qiyue Chen)"` 完整呈现且右侧边界无裁切。
        4. **附表排版与微生学物种空行规范 (Table 1 & Table 2 Integration):**
           - Table 1 与 Table 2 成功合并为 Master PDF 第 7 页（总页数严格锁定为 7 页）。
           - Table 2 严格遵循 Robin Seymour 标准 4 段横幅，Weekly Active Air 菌种各行之间严格保留空行且字号完全统一为 7.0pt (*Staphylococcus epidermidis*, *Pseudomonas parafulva*, *Bacillus safensis*, *Corynebacterium xerosis*, *Corynebacterium ureicelerivorans*)。
        5. **双文件法定交付铁律执行 (Strict 2-File Deliverable Protocol - Rule 0M):**
           - 严格交付 2 个核心文件：转录官方透视命名 Master 7-page PDF (`CORP-FORM-21- 01 Oct 2026 OOS-262289 Link Pharmacy (E19112) - ScanRDI.pdf`) 与 Standalone 1-page Word Table (`Tables OOS-262289 Link Pharmacy (E19112) - ScanRDI.docx`)，严禁提供任何冗余文件。



60. **Universal Font Size Uniformity & Zero Tampering on User-Pasted Content (复制粘贴零擅自改动字号与物种空行格式锁定 - 铁律闭环):**



    *   **核心准则 (Core Mandate):**

        - 在所有 OOS 案卷（ScanRDI, Celsis, USP <71>, EM）中，不论是向 Word 附表、主表单还是 PDF 表单填充、转录或复制粘贴内容时，**绝对禁止擅自调整、缩放或篡改任何字体大小（Font Size）与排版样式**。

        - 凡用户手动编辑确认的表格（如 Table 2 微生物鉴定 Microbial ID 各物种统一字号、物种间空行分隔），代码与转录流水线必须 100% 忠实保留原样格式，确保全表各行字体大小完全一致，空行层级分明。



61. **Master Table Golden Standards: Robin Seymour (QA) Mandates & User Core Requirements (OOS 附表双轨终极标准规范：Robin 质保法定规范与用户核心排版铁律 - 终极锁定):**

    *   **一、Robin Seymour (QA) 质保法定规范 (Robin's QA Golden Standards):**
        1. **4 大标准分段横幅用语铁律 (Canonical 4-Banner Dividers - 严格无日期后缀):**
           - Banner 1: `Personnel EM`（官方绝对标准：严禁任何日期后缀！仅在 QA 特殊要求的个别单日回溯案卷中允许带日期）。
           - Banner 2: `Biological Safety Cabinet EM for E00XXXX and E00YYYY`（官方绝对标准：严禁带 `(BSC)` 前缀，严禁重复出现 `BSC`，严禁日期后缀！多台 BSC 之间必须使用 `and`，严禁 `&`）。
           - Banner 3: `Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs`（L-Suite 必须完整列出所有洁净室编号 `145, 144, 143, 142`，或 `Suite 116 (E001738)`，严禁日期后缀；套间代号统一精简为 `(L-Suite)`）。
           - Banner 4: `Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs`（严格无日期后缀）。
        2. **连接词严禁使用 "&" (Strict Prohibition of "&"):**
           - 在横幅标题、设备列表及表内文本中，一律使用标准英文单词 `and`（如 `BSC E001938 and E001937`），绝对严禁使用非正式符号 `&`。
        3. **标准观察结果与注释规范 (Observation & Notes Capitalization Standard):**
           - 阴性结果统一使用 Title Case：`No Growth`（严禁全小写 `no growth` 或首词大写 `No growth`）。
           - 无菌或不适用列统一使用：`Not Applicable`（表格空间允许下严禁简写为 N/A）。
           - 正常无异常注释统一使用：`None`。
           - 异常情况说明首字母大写且末尾不加句号（如 `Plate was desiccated on 5 day read`）。
        4. **微生物双名法斜体铁律 (Binomial Nomenclature Italicization):**
           - 凡涉及属名与种名的微生物学名，必须严格使用斜体（*Italic*），如 *Staphylococcus epidermidis*, *Pseudomonas parafulva*, *Bacillus safensis* 等；属名首字母大写，种名全小写。
        5. **附表严格锁定 1 单页 (Strict 1-Page Layout Invariant):**
           - 独立 Word 附表（Table 1 + Table 2）经排版与 Word COM 导出为 PDF 时，总页数**必须严格锁定为 1 单页**，严禁溢出到第 2 页！合并入主案卷时作为 Page 7 挂载。

    *   **二、用户核心排版与数据要求 (User's Core Mandates):**
        1. **有 Microbial ID 则免除染色形态描述 (Microbial ID Precedence - 确证物种免染色形态):**
           - **“有Microbial ID就不用staining了”**：当送检 EM 平板已获得确证的微生物鉴定结果（Microbial ID 出具具体菌种名）时，Microbial ID 单元格内**直接填报确认的微生物学名，严禁再填充初步革兰氏染色形态**（如无需再写 `Gram (+) cocci` 或 `Gram (-) rods`），避免信息冗余与画蛇添足！
        2. **多菌种物种之间强制独立空行 (Mandatory Species Empty Line Separation):**
           - **“物种要空行”**：当单个 EM 样品检出多种微生物（如 ETX-261012-0433 检出 5 种菌）时，各菌种学名之间**必须强制插入独立空行 (`

`) 进行视觉分隔**，保持层级呼吸感，绝对严禁压缩合并或紧密挨合在单行内！
        3. **全表数据单元格字号绝对均一 (Strict Uniform Font Size - 7.0pt / 14 Half-Points):**
           - **“不过最好字体是都一样大小”**：全表（Table 1 样品信息与 Table 2 环境监测）所有数据单元格——包括 Sampling Site, Frequency, Date, Analyst, Observation, EM Plate ETX Number, Microbial ID, Notes 及超链接文字——其字体大小必须**严格保持绝对均一的 7.0pt (14 half-points)**！
           - 🚨 **底层 XML 超链接防坑铁律**：Word 中的超链接 `<w:hyperlink>` 拥有独立的 `<w:r>` 节点，常规遍历 `p.runs` 会漏改其字号（导致残留 6.5pt / `w:sz w:val="13"`）。必须在底层 XML 级别将所有 `<w:sz>` 与 `<w:szCs>` 统一锁死为 `14`（7.0pt），彻底消除局部字号差异！
        4. **原生可点击超链接嵌入 (Native Hyperlink Embedding):**
           - Table 1 中的主 Sample ID 与 Table 2 中所有检出菌落的 EM 平板 ETX 编号，必须嵌入原生 OpenXML `<w:hyperlink>`，直达 EagleTrax 电子送检详情页。
        5. **双文件交付法定归宿 (Rule 0M Compliance):**
           - 附表最终交付形式严格限定为 Word 版本：`Tables OOS-<ID> <Client> - <Test>.docx`。
           - 严禁单独交付独立的 Table PDF，其导出的 PDF 仅用于作为 Page 7 合并进完整主 PDF 案卷中。


62. **Good Documentation Practices (GDP) N/A Blank Field Sign-Off Standard (N/A + 分析员官方缩写 + 2位年份日期 DDMMMYY 规范):**
    *   **适用范围**: 全局适用于所有 OOS 调查模块（ScanRDI, Celsis, USP <71>, Environmental Monitoring）及官方表单（Phase I CORP-FORM-21, Phase II CORP-FORM-22 等）。凡任一文本框、栏目、表格单元格或溢出续页（如 Page 3 Section C "Other: ________" Text Field48、Page 5 "Phase I Summary (continued):" Text Field51 等）无需填写、不适用或划消签注时，必须严格执行该规范。
    *   **法定格式**: `N/A <Analyst Initials> <DDMMMYY>`（例如：`N/A QYC 21Oct26`）。
    *   **2 位年份绝对铁律（Strict 2-Digit Year Mandate）**: 日期部分必须且只能采用 2 位年份（如 `21Oct26`、`01Oct26`、`15Sep26`），**🚨 绝对严禁使用 4 位年份（严禁写成 `21Oct2026`）**！
    *   **分析员官方花名册对应（Analyst Initials Compliance）**: 缩写必须严格对应官方权威花名册（如 Qiyue Chen 必须为 `QYC`，严禁写成 `QC` 或 `KC`，遵守 RULE 0A）。
    *   **空格与标点铁律**: 各字段间以单个空格分隔，日期中不得包含斜杠或连字符，句尾严禁加句号。


63. **Robin Seymour (QA Director) Official Review Golden Standards for USP <71> Phase I & Phase II (Solyn LLC OOS-262098 质保复核定稿标杆与六大黄金准则 - 终极锁定):**
    *   **标杆案卷来源**: 基于 QA Director Robin Seymour (RS) 审阅并签署的官方标杆报告 `OOS-262098 Solyn LLC (E75000) - USP71 -RS Reviewed` 及 `... - Phase II - RS Reviewed`，提炼并固化以下六大质保法定准则：
    *   **1. 设备栏引用与生物安全柜强制罗列 (Equipment Section "Phase I Summary" & Mandatory ISO 5 BSC):**
        - 在 `CORP-FORM-21` 设备栏目中，若环境监控与控制引用正文叙述，必须规范写为 `"Phase I Summary"`（🚨 **严禁加 "See"，严禁写成 "See Phase I Summary"**）。
        - 在设备列表中，除了培养箱与传感器外，必须明确列出主处理生物安全柜 `ISO 5 BSC E001316` 及其校准到期日 `Jun 2027`（与 Phase II 绝对对齐，经 QA Tahsin Kevenk 审核定稿）。
    *   **2. 冻干粉/小容量注射剂复溶重构过程详述铁律 (Mandatory Reconstitution Process Narrative):**
        - 在 Phase I 样品处理段落（Page 3 `Text Field49`）中，凡待测样品为冻干粉或小容量注射剂（如 GLP3R/Cagrilinitide 25mg/5mg），**必须详细写明复溶溶媒、体积、浓度及过滤分液比例**：
          `"...reconstituted with 2mL Bacteriostatic Water per vial (each vial containing GLP3R/Cagrilinitide in 25mg/5mg concentration). Half the contents of each reconstituted vial was then dispensed directly from the container into sterile canisters placed on a vacuum manifold for membrane filtration."`
          严禁仅笼统写 `"45 gm filtered per media across 3 vials"`。
    *   **3. 洁净室套间 115 认证到期日统一标准 (CR115 Certification Due Date Standard):**
        - 洁净室套间 115（`CR 115 / E001737`）的校准认证到期日统一为 **`Dec 2026`**（与 CR 114 一致），严禁写为 `June 2027`。
    *   **4. 先前同批快速法合格与当前药典法不合格冲突时的质保官方两段式防守金标准 (RS Golden 2-Paragraph Defense for Prior Passing Rapid Test):**
        - 当同一批号此前送检快速法（如 ScanRDI）合格，但 USP <71> 药典法定检验与复测均不合格时，必须严格执行 Robin Seymour 官方审阅定稿的两段式防守措辞：
          - **第 1 段 (方法适用性缺失防守)**:
            `"It was identified that the same sample lot (2608-216) had previously been submitted by the client for ScanRDI sterility testing under submission ETX-260807-0602 and yielded passing results. However, it is important to note that a review of the submission determined that no approved method suitability test was available to support the validity for microbial growth recovery using the ScanRDI method for this specific sample."`
          - **第 2 段 (低浓度微生物非均一分布防守与不可比定论)**:
            `"Furthermore, the discrepancy between the previous passing ScanRDI result and the current USP <71> sterility failure may be attributable to the non-homogeneous distribution of microbial contamination within the compounded batch. When microbial contamination is present at very low levels, microorganisms may not be uniformly distributed throughout the product, resulting in variability between samples collected from the same lot. Consequently, one aliquot may not contain contaminating microorganisms and yield an acceptable result, while another aliquot from the same batch may contain viable microorganisms and generate a positive sterility test result. Given the absence of a qualified method suitability study for the ScanRDI submission, the prior passing result cannot be considered directly comparable to the current USP <71> sterility test outcome."`
    *   **5. Phase II 根因分析极简法定用语 (Phase II Root Cause Statement Simplification):**
        - 在 `CORP-FORM-22` 第 5 页 Section D 根因陈述（`Text Field25`）中，**严禁堆砌长篇恢复动力学或 ScanRDI 讨论**，必须聚焦于复测确证同种微生物的核心事实，精简为两句话：
          `"Inherent sample contamination within Lot 2608-216, supported by recovery of the same organism (Microbacterium sp. PM5) during retesting. The original USP <71> sterility failure was confirmed as valid."`
    *   **6. 调查结论双检验有效性法定句式 (Investigation Conclusion Validating Both Tests):**
        - 在 Phase II 调查结语中，必须明确陈述批次污染并判定**初始测试与复测均为有效失败**：
          `"Therefore, inherent product contamination in Lot 2608-216 is confirmed, both the initial failing result for ETX-260902-0505 and the retest ETX-260914-0470 is considered valid."`
    *   **7. Section E 主管与质保处置权定论 (Section E QA Exclusive Jurisdiction):**
        - `CORP-FORM-22` Section E 属于质保主管专属签署权限，起草人必须完全留空。Robin Seymour 亲自在两个文本框中签注为：`"N/A RS 21 Oct 2026"`。


64. **Three-Tier OOS Review, Roles & Signature Hierarchy Protocol (OOS 报告三级签名流转法定层级与角色绝对清晰铁律 - 🚨 终极锁定):**
    *   **根本病灶与角色认知纠正**: 严禁将直接实验主管（Microbiology Supervisor）与独立的质保部（Quality Assurance / QA）混为一谈！在无菌微生物实验室，OOS 调查报告签署与审核有着严格的层级与部门分工：
        1. **第一级：起草与调查人 (Prepared by: / Initiating Analyst)**:
           - **人员**: **Qiyue Chen (QYC)**（Microbiologist / 调查报告起草人）。
           - **职责**: 负责穿透调取实验原始记录、环境监测数据、复溶过滤参数，起草完整调查报告，并签署 Prepared by: 槽位（Signature_0）。
        2. **第二级：实验室直接主管 (Reviewed by: / Microbiology Supervisor)**:
           - **人员**: **Robin Seymour (RS)**（Microbiology Lab Supervisor / 直接主管）。
           - **职责**: 负责对调查员提交的报告进行第一道技术与实验室复核（Supervisor Review），提出修改意见、处置 Section E（填写 N/A RS ...），在修改定稿后由调查员先签名，再由主管在 Reviewed by: 槽位签名（Signature_1）。
        3. **第三级：独立质保部最终批准 (Approved by (QA): / Quality Assurance Department)**:
           - **部门/人员**: **Quality Assurance (QA)**（独立的质保部门，如 QA Reviewer / Director）。
           - **职责**: 实验室内（调查员与主管）联签完成之后，报告才正式流转至 QA 进行最终的合规审查与结案签署（Approved by (QA): / Signature_2）。**在实验主管审核期间，QA 尚未介入检查！**
    *   **流转路线绝对铁律**:
        Qiyue Chen (Prepared by) ✍️ 签署 -> 流转至 Robin Seymour (Reviewed by) ✍️ 审核签署 -> 流转至 QA 质保部 (Approved by) ✍️ 终审批准。



62. **Sterile Suite & Biosafety Cabinet Analyst Assignment Ground Truth (无菌套间、洁净室与分析员常态化主工位权威映射 - 🚨 终极锁定):**

    *   **业务背景与权威实战花名册 (Laboratory Operational Ground Truth)**:
        根据无菌检验（Celsis, Scan RDI, USP <71>）在 Eagle Analytical 实验室现场的真实日常排班与操作工位分布，各分析员在各洁净室及主/副生物安全柜之间的常态化分配规则严格守恒如下：
        
        1. **Suite 114 常态化排班与工位**:
           - **Cleanroom 114B (ISO 7 核心室 / 主操作 BSC E001316)**:
             - 专职样品初次接种与过滤处理（Processing Analyst）。
             - 常态化操作分析员：**ES (Alex Saravia)**、**GS (Gabrielle Surber)**。
           - **Cleanroom 114A (ISO 7 缓冲室 / 分装与换机 BSC E001798)**:
             - 专职培养 7 天后取样分装（Aliquoting Analyst）及近机操作（Celsis Advance 2 E002222）。
             - 常态化操作分析员：**CCD (Cuong Du)**、**ALA (America Alanis)**。
        
        2. **Suite 115 常态化排班与工位**:
           - **Cleanroom 115B (ISO 7 核心室 / 主操作 BSC E001309 / E001312)**:
             - 专职大批量无菌样品过滤与经典培养处理。
             - 常态化操作分析员：**AOD (Ayomide Odugbesi)**、**GL (Guanchen Li)**。
        
        3. **L-Suite (Suite 145 / 144) 常态化排班与工位 (RS Verified Ground Truth)**:
           - **Cleanroom 145 (`CR145 / E001979`, Cert Due `Nov 2026` — ISO 7 核心室 / Scan RDI 处理 BSC `E001938`, Cal Due `Jul 2027`)**:
             - 专职样品引入、预贴签与过滤处理（Processing Analyst）。
           - **Cleanroom 144 (`CR144 / E001978`, Cert Due `Nov 2026` — ISO 7 缓冲室 / Scan RDI 换台清场 BSC `E001937`, Cal Due `Jul 2027`)**:
             - 专职换批消毒清场（Changeover Analyst）。
           - 常态化操作分析员：**ELB (Ellen Bennett)**、**KSM (Karla Silva)**、**VV (Varsha Subramanian)**、**EN (Elysse Nioupin)**。




*   **Mandatory Raw EM Logbook Ground-Truth Audit & Positive Plate Lifecycle Protocol (Rule 46 / Rule 65):**

 Mandatory Raw EM Logbook Ground-Truth Audit, Anti-Template Hallucination & Positive Plate Lifecycle Protocol (环境监测原始扫描件地毯式核查、反模板脑补与阳性平板生命周期法定铁律 - 终身基准)

**CRITICAL**: When investigating OOS cases and generating environmental monitoring tables (Table 2 & Table 3 across ScanRDI, Celsis, USP <71>, and EM):

1. **绝对禁止依赖模板惯性与假数据脑补 (Strict Ban on Template Inertia & Hallucinated Defaults)**:
   - 当用户上传或提供环境监测纸质扫描件（如 `svc-scan@...` 或包含 F01/F02/F03/F06/F10 的 PDF）时，**严禁使用千篇一律的通用模板假数据**（如盲目填入 `02Oct26 ES 0 CFU`、`09Oct26 CCD 0 CFU`、`No Growth`、`None`）。
   - 必须通过图像处理自动校正旋转角度（如将 270°/90° 倒立扫描件恢复至正立），逐页对齐核查每一张台账原件。

2. **真实当班日期与物理设备历史锚定法则 (Raw Bench Shift & Hood Precedence Rule)**:
   - **人员监测前序日 (Personnel EM Date Before Testing)**：必须严格锚定台账中该分析员**真实记录的上一当班日期**（例如 ES 接种前台账记录为 `29Sep26`，则必须填 `29Sep26`，绝对严禁按周五推算假定不存在的 `02Oct26`！）。
   - **生物安全柜前序日 (BSC EM Date Before Testing)**：严格跟踪该台生物安全柜的真实运行历史。若前序日由其他人员使用（例如 `01Oct26 QYC` 在 BSC 1316 执行 Practice），必须真实记录 `01Oct26` `QYC` `No Growth`，严禁擅自抹杀或偷换为样品分析员！
   - **分段横幅严格对应单设备 (Single BSC Divider per Table)**：Table 2 仅写初始接种主设备（如 `Biological Safety Cabinet EM for E001316`）；Table 3 仅写分装主设备（如 `Biological Safety Cabinet EM for E001798`），严禁偷懒合并双设备。

3. **动态浮游菌检出与送检单号生命周期铁律 (Active Air Recovery & Positive Plate Lifecycle Rule)**:
   - 凡浮游菌或表面出现菌落检出（如 29Sep26 在 ISO 8 114 检出 3 CFU）：
     - **Observation**：严格采用复数规范格式（如 **`3 CFUs`**）。
     - **EM Plate ETX Number**：完整提取台账圈注的手写送检单号（如 **`ETX-261006-0660`**），并嵌入标准可点击超链接。
     - **Microbial ID**：只要已登记送检编号且最终报告尚未出具，**状态必须严格填入 `Pending`**！**绝对严禁写成 `None`！**（写 None 会让 QA 误判为未长菌或未送检，构成严重 cGMP 事实偏差）。
     - **Notes**：保持极简统一，设为 **`None`**（常规合格项无需赘述行动限）。

4. **强制自检与证据先导门禁 (Pre-Response Forensic Evidence Gate)**:
   - 严禁在未真正完成全量切片阅读前口头宣称“已阅读扫描件”。向用户提交交付物时，必须主动汇报关键页码的真实发现（如检出 CFU、送检单号、特殊操作者签字等），确保每一项数据 100% 经得起原始台账核验。


70. **Robin Seymour QA Review Standards & Facility Layout Ground Truth (RS Reviewed 终极审核规范与洁净室布局真理 - 🚨 终极锁定):**
    *   **Robin Seymour (RS) 主管地位与执行铁律 (Supervisor Authority Mandate):**
        - **组织关系**: **Robin Seymour (RS)** 是报告撰写人 **Qiyue Chen (QYC)** 的直接实验室主管（Microbiology Laboratory Supervisor）。
        - **最高执行准则**: RS 在实验室管理、技术规范、套间真实布局映射、理化证据链审查及官方审核意见上具备最高权威。在所有 OOS 调查案卷中，我们**必须 100% 严格遵从 RS 的要求与审核修改**，绝不允许擅自变动或以旧模板覆盖！
    *   **ScanRDI 事件总数必须核实原始数据包 (Event Count Source Truth):**
        - 严禁仅依赖分析员手写便签或口头流转数据（Analyst data entries）。**必须严格调取 Scan Data Packet 仪器原始打印报告核验总事件数（Total Events）与确证微生物事件数（Confirmed Microbial Events）**。
        - 例如 OOS-262289 中分析员记录为 15，实际原始数据包显示为 19 total events, 4 confirmed microbial events，以原始数据包 19 为唯一法定依据。
        - 🚨 **用户协作与主动提问门禁 (User Collaboration Gate)**: Scan Data Packet 原始打印件在系统或扫描附件中往往难以自动定位。Agent **绝不可自行盲目采信手写便签或胡乱推测，必须主动向用户求助提问**（如：*“Scan Data Packet 原始数据包较难定位，请帮我核验本次检验的真实 Total Events 与 Confirmed Microbial Events”*），由用户协助查找并提供，确保数据 100% 真实准确！
    *   **L-Suite 洁净室与生物安全柜真实物理布局映射 (L-Suite Physical Layout Mapping):**
        - **BSC E001938** 真实物理位置位于 **Cleanroom 145 (CR145)**（用于样品引入、预贴签与过滤处理）。
        - **BSC E001937** 真实物理位置位于 **Cleanroom 144 (CR144)**（用于换批消毒清场 Changeover）。
        - 🚨 **严禁混为一谈**：绝不可将两台 BSC 笼统写在 CR144！
    *   **洁净室与 BSC 认证设备及校准期精准填报 (Equipment Certification Ground Truth):**
        - Cleanroom 认证设备栏（Page 2 `Text Field32`）涉及 L-Suite 时，必须完整罗列两个操作套间：第一行 `E001978 (CR144)`，第二行 `E001979 (CR145)`（字号 5.0pt）。
        - 对应校准有效期（`Text Field33`）必须精准对应（如两者均为 `Nov 2026\rNov 2026`）。
        - BSC 校准有效期（`Text Field44`）：BSC E001938 与 BSC E001937 校准期统一核定为 `Jul 2027`。
        - ScanRDI 仪器校准期（`Text Field31` & `Text Field35`）：E001040 实际校准期为 `Oct 2026`（非 Nov 2026）。
    *   **环境菌形态学对比的科学防御性原则 (Scientific Defensibility on Morphology Comparisons):**
        - 🚨 **质保重大防线**：当环境监测检出菌（如 *Staphylococcus epidermidis*）在微生物学形态分类上同样属于球菌（Cocci）时，**绝对严禁草率断言“环境检出菌形态与阳性样本 Cocci 形态截然不同（distinctly different）”**，否则会被审计专家或质保直接判定为科学常识性逻辑漏洞！
        - **Robin 质保法定处理标准**：直接完全删除形态学对比陈述，将防守重心完全聚焦在**物理空间隔离、闭门操作、向外逐级正压气流屏障，以及关键 ISO 5 操作区台面/沉降碟/手套完全无菌**的核心证据链上。
    *   **月度清洁消毒法定精炼用语 (Monthly Cleaning Standard Statement):**
        - 统一采纳 Robin 质保规范句式：
          `"It was documented that all chemical indicators passed, which confirmed the efficient monthly cleaning of all cleanrooms in L-Suite."`
    *   **同日其他阳性样本极简披露标准 (Streamlined Concurrent Positive Disclosure):**
        - 无需披露第三方客户名称与具体产品规格，保持精炼客观合规：
          `"One other sample (ETX-XXXXXX-XXXX, OOS-XXXXXX) yielded a positive result on [DDMMMYY]. However, that sample was processed by a different processing analyst during a different shift. The two testing events were segregated by shift, processing analyst and primary containment equipment. Furthermore, analyst [Name] verified that gloves were disinfected between samples and standard aseptic techniques were strictly maintained. Consequently, there is no evidence of sample-to-sample cross-contamination or a shared laboratory vector."`


## Pending/Future Work



*   **Roll out Smart Justification to USP <71>:** The engine is live for Celsis and Scan RDI, but `USP71.py` still needs its underlying logic updated to utilize the 4-Step Shielding Mechanism and the new "RS Reviewed" narrative format (adjusting for its specific workflow).



*   **Template Updates:** The underlying `.docx` and `.pdf` templates need manual layout updates (by the user in Word/Acrobat) to accommodate the significantly longer narrative text before they can be perfectly auto-filled without cutoff/font-shrinking.



*   **OOS Table Generator (`generate_oos_table.py`) — Phase 2:** Core script is built and verified for all 4 OOS types (EM, Scan, Celsis, USP71). Phase 2 will standardize the AI extraction prompt/checklist for reading scanned PDF logbook images and outputting structured data dicts. The workflow is: user provides scanned PDF copy → AI reads images and extracts ETX IDs, CFU counts, dates, analysts → AI builds data dict → script fills existing `tables for XX.docx` templates → outputs final Word table. Microbial IDs provided separately by user.





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


66. **OOS-262329 GoGoMeds Select (E10747) - USP <71> BUD-T90 Investigation Standards (🚨 终极锁定):**
    *   **案卷基本信息**:
        - **OOS ID**: `262329`
        - **Client**: `GoGoMeds Select (E10747)`
        - **Sample ID**: `ETX-260624-0251 (BUD-T90)`
        - **Analyte**: `Sermorelin 0.9mg/ml` (Lot: `17009`, Container size: `5 mL`, Batch: `500`)
        - **Test Method**: `USP <71> / EP 2.6.1 Sterility Test` [T0001006726]
        - **Technique**: Direct Inoculation (Approved Method Suitability `ETX-260401-0599` on file: `15mL tween80 + 5mL sample into 200mL media`)
        - **Incident Date**: `12-Oct-2026` (Day 18 read by Andrew Carrillo; status changed to OOS-262329)
        - **Disposition**: Require Retest - Phase II (CORP-FORM-22 required for stability study)
    *   **核心实验数据与法定理化证据链**:
        - **接种与培养**: 由 `EN` 于 22Sep26 准备样品，`GL` 于 23Sep26 在 Suite 115 ISO 5 BSC E001314 完成直接接种并于 24Sep26 12:55 AM 入箱培养。
        - **混浊分装与动力学**: Day 14 (08Oct26) 因高浓度 Sermorelin 样品本身基质混浊判定为 Inconclusive，由 AOD 于 07Oct26 6:59 PM 转移分装至新鲜培养基 (-46/54-A 及 -47/54-A)。Day 18 (12Oct26) 原样瓶 (-46/54) 与分装瓶 (-46/54-A) 之 TSB 均呈阳性生长，FTM 原样瓶及分装瓶全程阴性澄清。
        - **同批次与交叉污染排除 (Sibling Stability Sample Ground Truth)**: 接种员在 23Sep26 同批次处理约 50 个样品，包含同客户另一稳定性样品 ETX-260624-0206 (T90, Sermorelin 3 mg/mL, Lot 17008)，在同等条件下操作且结果均为阴性；且 AOD 于 07Oct26 分装时所有同步分装样品均无生长，彻底排除批次操作失误与样品间交叉污染。
        - **微生物鉴定**: 送检单号 `ETX-261012-0367`，经差速染色（革兰氏阳性杆菌）及确证鉴定为 ***Bacillus megaterium***。
        - **环境监测 (EM - 14 页台账地毯式核查与 Rule 46 真实设备运行历史锁定)**:
          - **人员手套 (Personnel EM)**: 接种员 GL 连续 3 日（22Sep26 BSC 1312, 23Sep26 BSC 1314, 24Sep26 BSC 1312）左右手接触平皿均为 0 CFU，全为 **No Growth**。
          - **生物安全柜 (BSC E001314)**: 严格跟踪物理设备真实运行历史：
            - 前序日 (22Sep26): 由 AOD 运行 71 Sterility，表面 (4 点) 与沉降 (2 点) 均为 0 CFU (**No Growth**)。
            - 检验日 (23Sep26): 由 GL 运行直接接种，表面 (4 点) 与沉降 (2 点) 均为 0 CFU (**No Growth**)。
            - 后续日 (24Sep26): 由 AOD 运行 71/Celsis，表面 (4 点) 与沉降 (2 点) 均为 0 CFU (**No Growth**)。
          - **Suite 115 洁净室周表面监测**: 23Sep26 (检验当周周三) 由 SMO 监测 6 个位点，全为 0 CFU (**No Growth**)。
          - **Suite 115 洁净室周浮游菌监测**: 23Sep26 由 SMO 监测，核心 ISO 7 115A 与 115B 均为 0 CFU (**No Growth**)；最外层 ISO 8 Anteroom 115 检出 **6 CFUs**，送检微生物鉴定单号为 ETX-261001-0625，状态严格依规填入 **Pending**。按 Rule 0I 级联正压差与消毒封闭转运排除污染。
        - **月度清洁**: Suite 115 月度清洁由 `EN`, `SAA`, `ANC` 于 `13Sep26` 执行，H2O2 指示卡（Lot `260620`, Exp `20Jun28`）全项合格。
        - **6 个月历史**: 官方追踪台账显示 GoGoMeds Select 之 Sermorelin 0.9mg/ml 在 USP <71> 下无前科（Zero Prior Failures）。
    *   **双文件交付物规范 (Rule 0M)**:
        - **文件 1 (Master Transcribed PDF)**: `CORP-FORM-21- 12 Oct 2026 OOS-262329 GoGoMeds Select (E10747) - USP71.pdf`（7 页完整合并版，含 6 页 AcroForm 表单及 Page 7 Table 1 & 2）。
        - **文件 2 (Word 版独立附表)**: `Tables OOS-262329 GoGoMeds Select (E10747) - USP71.docx`（严格 1 单页版）。
        - 唯一定位落盘至企业 OneDrive 桌面：`C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\`。


67. **Weekly Environmental Monitoring (EM) "Week of Testing" Calendar Alignment Protocol (周度环境监测 "Week of Testing" 自然周严格对齐铁律 - 🚨 终极锁定):**
    *   **核心法则**: 在所有无菌检验调查（Celsis、ScanRDI、USP <71>）附表 Table 2 & Table 3 中：
        - `Week of Testing`（检验当周）必须严格与测试日期所在的同一个 Monday-to-Sunday 自然周对齐！
        - **严禁整周错位**：若检验在周一进行，当周的周检取当周中（如周四）采样的记录，严禁将上一周（T-1 周）的周检记录错填为检验当周。
        - **官方先例基准**:
          - `OOS-262017`: 24Aug26 (Mon) Processing ➡️ Table 2 为 25Aug26 (Tue)；31Aug26 (Mon) Aliquoting ➡️ Table 3 为 04Sep26 (Fri)。
          - `OOS-262333`: 05Oct26 (Mon) Processing ➡️ Table 2 为 08Oct26 (Thu)；12Oct26 (Mon) Aliquoting ➡️ Table 3 为 15Oct26 (Thu)。

68. **OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis Sterility Investigation Standards (🚨 终极锁定):**
    *   **案号基本信息**:
        - **OOS ID**: `262333`
        - **Client**: `Optimal Balance Pharmacy (E19193)`
        - **Sample ID**: `ETX-261002-0288` (Subculture: `ETX-261012-0654`)
        - **Analyte**: `TESAMORELIN 8 MG/ML (3 ML) Injection` (Lot: `LG342011076`, Container: 3 mL vials, Total 10 vials inoculated = 15 mL sample + 3 mL Tween 80 per media jar)
        - **Test Method**: `Celsis Sterility Test` (`MICRO-SOP-44` Rev 13 / `MICRO-SOP-53` Rev 04)
        - **Approved Method Suitability**: `ETX-250826-0390`
        - **Incident / Readout Date**: `12-Oct-2026`
        - **Initiator / Written By**: Initiator: `Cuong Du` (Written by `Qiyue Chen`)
        - **Personnel Roster**: Prepping: `Mukyung Jang (MJ)`, Processing: `Alex Saravia (ES)`, Aliquoting: `Cuong Du (CCD)`, Instrument Reader: `Cuong Du (CCD)`
        - **Disposition**: Laboratory Investigation Phase I Complete; No assignable laboratory error; Original test result valid.
    *   **核心实验数据与法定理化证据链**:
        - **接种与培养**: 02Oct26 样品到样冷藏，05Oct26 由 `ES` 在 Suite 114 ISO 7 Cleanroom 114B 内 ISO 5 BSC E001316 完成直接接种（1 x 300mL FTM + 1 x 300mL TSB，各加 3mL Tween 80），匹配阴性对照同步操作，入箱培养。
        - **分装与检测**: 12Oct26 培养期满，由 `CCD` 在 ISO 7 Buffer Room 114A 内 ISO 5 BSC E001798 执行无菌分装；使用 Advance 2（`E002222`，Workload Report: `101226-2222-5.wkl`）完成 Celsis ATP 读板检测。
        - **读板数值**:
          - FTM (Bottle 1/5): **Positive, 3,647,568 RLU** (Positive Cut-off: 2,596.5 RLU)
          - TSB (Bottle 4/5): **Negative, 29 RLU** (Positive Cut-off: 2,326.5 RLU)
          - 系统对照全合格: Instrument Blank 14 RLU (限度 ≤100), Reagent Blank 64 RLU (限度 ≤1,000), ATP Positive Control 83,684 RLU (限度 ≥50,000, 批号 022601-1483, 效期 Jan 2027), 阴性对照 FTM 866 RLU / TSB 776 RLU。
        - **微生物确证鉴定**: 差速染色为革兰氏阳性芽孢杆菌（Gram-positive rods, spore-forming），三代 Oxford Nanopore 测序比对鉴定为 ***Bacillus paralicheniformis*** (85.73% match)。
        - **环境监测数据链 (Rule 0V 自然周严格对齐)**:
          - Processing 周 (05Oct26 Mon): 当周周检为 **08Oct26 (Thu)** by SMO（P20 表面采样 0 CFU, P22 浮游菌采样 0 CFU，各点位全为 **No Growth**）。
          - Aliquoting 周 (12Oct26 Mon): 当周周检为 **14Oct26 (Wed)** by SMO（P21 表面采样 0 CFU 全为 **No Growth**；P23 浮游菌采样：ISO 7 114A 与 ISO 7 114B 均为 **0 CFU (No Growth)**，最外层 ISO 8 Anteroom 114 检出 **6 CFUs**，送检微生物鉴定单号为 **`ETX-261021-0578`**，鉴定状态为 `Pending`）。
          - 原 29Sep26（含 ISO 8 3 CFU）属于前序周已依法移出。14Oct26 ISO 8 Anteroom 之 6 CFUs 依据 Rule 0I 确认为外围更衣缓冲区域，且样本均经封闭消毒转运，层级压差与物理密闭门彻底杜绝逆向扩散，所有 ISO 7 洁净室与 ISO 5 操作区 100% 无菌受控（No Growth）。
          - 每日人员 (ES 05Oct, CCD 12Oct) 与 BSC (E001316, E001798) 表面/沉降全为 0 CFU (No Growth)。
        - **月度清洁**: Suite 114 月度清洁由 `SAA` 于 `27Sep26` 执行，HaloMist (Lot 2606A, Exp 06Jun28)，过氧化氢指示卡 (Lot 260620, Exp 20Jun28) 全检合格。
        - **6 个月特定分析物历史**: 依据官方 SharePoint 台账，Optimal Balance Pharmacy 之该特定分析物 "TESAMORELIN 8 MG/ML (3 ML) Injection" 在 Celsis 无菌检验方法下无前科（Zero Prior Failures）。
        - **交叉污染评估**: 检验当日处理的其他所有样品结果均为阴性，排除样品间交叉污染。
    *   **双文件交付物规范 (Rule 0M)**:
        - **文件 1 (Master Transcribed PDF)**: `CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf`（8 页完整合并版，含 6 页 AcroForm 官方表单及末尾 Page 7 Table 1 & 2，Page 8 Table 3）。
        - **文件 2 (Word 版独立附表)**: `Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx`（严格 2 页版）。
        - 唯一法定落盘物理路径：企业 OneDrive 桌面 `C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\`。

69. **Global Dual-Track EM Inspection & Growth Focus Screenshot Collaboration Protocol (全局环境监测双轨排查与菌落生长聚焦截图法定协作规范 - 🚨 终极锁定):**
    *   **核心法则**: 在所有 OOS 调查案卷（ScanRDI、Celsis、USP <71>、EM）中，全面推行“全景归档 + 焦点透视”双轨协作模式：
        1. **全景后台拉网 (Agent 主导)**：用户将全量多页 PDF 归档至 EagleTrax；Agent 自动下载并拉网式提取常规阴性数据、设备到期日、物料批号效期、月度清洁消毒记录及人员初始对应。
        2. **异常焦点聚焦 (用户截图主导)**：凡涉及平皿生长（Growth / CFUs > 0）、数字划改更正或手写子送检单号（Subculture ETX ID）的页面，用户直接在对话框提供单页高清截图或一句话摘要（如 *“P23 ISO 8 浮游菌有 6 CFU，单号 ETX-261021-0578”*）。
        3. **精准锚定与防守撰写**：Agent 将截图作为法定理化证据锚点，提取真实 CFU 数、读板人与时间、子送检单号并设为 Pending；在 Page 5 自动整合 Rule 0I 级联压差与物理隔离防守论证。
        4. **零脑补门禁**：若截图字迹仍有模糊，Agent 必须停步提问，绝不臆测。


---

### Rule 71: Official L-Suite Cleanroom & Equipment Layout Ground Truth (Figure 6 - H2O2 Strips & Equipment Mapping)

**CRITICAL**: All cleanroom investigations, narratives, and equipment allocations involving **L-Suite** MUST strictly adhere to the official facility floor plan (**Figure 6: Locations for H2O2 Strips in “L” Suite**):

1. **L-Suite Overall Architecture & Hierarchy (From Outermost to Core)**:
   - **Room 142 (ISO 8 Entrance Room / Airlock)**: Outermost transition room containing **Sink** and **Cart**. Provides first-line physical barrier and pressure boundary.
   - **Room 143 (ISO 8 Anteroom & Instrument Room)**:
     - **Gowning Area**: Dedicated gowning zone with Cart.
     - **Section I**: Contains **Scan E002017**, central **Table**, **Incubator E001032**, and **Scan E001511**.
     - **Section II**: Contains **Scan E001230** (left) and **Scan E001040** (right, used for ScanRDI reading).
   - **Room 144 (ISO 7 Buffer Room / Cleanroom)**:
     - **Section I**: Contains **BSC E001988** (horizontal, top wall) and **BSC E001937** (vertical, right wall, used for Changeover).
     - **Section II**: Contains **Incubator E001034** (bottom-right) and Carts near entry door from Room 143.
   - **Room 145 (ISO 7 Processing Core Cleanroom)**:
     - **Section I**: Contains **BSC E001938** (horizontal, top wall, used for sample introduction, pre-labeling, and filtration) and **BSC E001319** (vertical, left wall).
     - **Section II**: Contains **BSC E001317** (horizontal, top wall) and mobile carts.

2. **Official Biosafety Cabinet (BSC) Room & Section Mapping Ground Truth**:
   - 🚨 **Room 145 (ISO 7)**:
     - **`BSC E001938`**: Located in **Room 145 Section I** (Top wall) → ScanRDI Processing / Filtration.
     - **`BSC E001319`**: Located in **Room 145 Section I** (Left wall).
     - **`BSC E001317`**: Located in **Room 145 Section II** (Top wall).
   - 🚨 **Room 144 (ISO 7)**:
     - **`BSC E001937`**: Located in **Room 144 Section I** (Right wall) → ScanRDI Changeover.
     - **`BSC E001988`**: Located in **Room 144 Section I** (Top wall).
   - 🚨 **严禁搞混**: E001938 绝不在 144，E001937 绝不在 145！两台核心 BSC 分属不同物理洁净室（145 与 144），中间有套间墙体和门物理隔离。

3. **Official ScanRDI Instrument Mapping (Room 143, ISO 8)**:
   - All ScanRDI instruments are housed in **Room 143 (ISO 8)**:
     - **`Scan E001040`**: Room 143 Section II (Right).
     - **`Scan E001230`**: Room 143 Section II (Left).
     - **`Scan E002017`**: Room 143 Section I (Left).
     - **`Scan E001511`**: Room 143 Section I (Right).

4. **Monthly Cleaning H2O2 Indicator Strips 13-Point Mapping (MICRO-SOP-9 Ground Truth)**:
   - **Strip 1**: Room 145 Section I (NW corner)
   - **Strip 2**: Room 145 Section I (SW corner)
   - **Strip 3**: Room 145 Section II (NE corner)
   - **Strip 4**: Room 145 Section II (SE corner)
   - **Strip 5**: Inside **BSC E001319** (Room 145 Section I)
   - **Strip 6**: Inside **BSC E001938** (Room 145 Section I)
   - **Strip 7**: Inside **BSC E001317** (Room 145 Section II)
   - **Strip 8**: Room 144 Section I (NW corner)
   - **Strip 9**: Room 144 Section II (SW corner)
   - **Strip 10**: Room 144 Section I (NE corner)
   - **Strip 11**: Room 144 Section II (SE corner)
   - **Strip 12**: Inside **BSC E001988** (Room 144 Section I)
   - **Strip 13**: Inside **BSC E001937** (Room 144 Section I)


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
