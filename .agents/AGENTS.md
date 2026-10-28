# Project-Scoped Rules for Eagle Analytical OOS Tools

## 📌 LATEST CRITICAL UPDATES (PINNED AT TOP — NEVER TRUNCATED)
> **MANDATORY FOR ALL AGENTS**: Before generating or editing ANY OOS report or table, you MUST also read the full rules at the bottom of `.agents/AGENTS.md` (`StartLine=800` to `960`) and `PRD.md` (`StartLine=3000` to end). Key highlights locked from **Supervisor Robin Seymour (RS)** & Official Facility Layout (**Figure 6**):
> 1. **Supervisor Authority (`RULE 0X.0`)**: **Robin Seymour (RS)** is Qiyue Chen's direct supervisor (Microbiology Lab Supervisor). All RS review corrections and phrasing standards are absolute law.
> 2. **ScanRDI Event Count (`RULE 0X.1`)**: Must come from the raw **Scan Data Packet** (e.g., `19 total events, 4 confirmed microbial events`), NEVER from analyst EagleTrax data entries. **If you cannot locate the raw Scan Data Packet, STOP and ask the user to verify it!**
> 3. **Official L-Suite Floor Plan & Equipment Mapping (`RULE 0X.2` & `RULE 0Y` / Figure 6)**:
>    - **Room 145 (ISO 7, `E001979`, Cert Due `Nov 2026`)**:
>      - **Section I**: **`BSC E001938`** (Cal Due `Jul 2027`, Strip #6 — ScanRDI sample introduction, pre-labeling, filtration) and **`BSC E001319`** (Strip #5); Room Strips #1, #2.
>      - **Section II**: **`BSC E001317`** (Strip #7); Room Strips #3, #4.
>    - **Room 144 (ISO 7, `E001978`, Cert Due `Nov 2026`)**:
>      - **Section I**: **`BSC E001937`** (Cal Due `Jul 2027`, Strip #13 — ScanRDI Changeover) and **`BSC E001988`** (Strip #12); Room Strips #8, #10.
>      - **Section II**: **`Incubator E001034`**; Room Strips #9, #11.
>    - **Room 143 (ISO 8)**: Gowning Area + **Section I** (`Scan E002017`, `Table`, `Incubator E001032`, `Scan E001511`) + **Section II** (`Scan E001230`, `Scan E001040` [Cal Due `Oct 2026`]).
>    - **Room 142 (ISO 8)**: Outermost entrance room with `Sink` and `Cart`.
>    - **Page 2 Cleanroom Fields (`Text Field32`/`33`)**: When testing spans CR145 & CR144, list BOTH: `E001978 (CR144)\rE001979 (CR145)` (`5.0pt`) and `Nov 2026\rNov 2026` (`5.0pt`).
> 4. **Active Air Morphology Comparison Ban (`RULE 0X.4`)**: **NEVER** claim environmental isolates (e.g., *Staphylococcus epidermidis*) have "morphology distinctly different from cocci" (since *Staphylococcus* IS cocci-shaped!). Rely strictly on physical room/door segregation, positive pressure cascade, and zero ISO 5/glove growth.
> 5. **RS Streamlined Phrasing (`RULE 0X.5`)**:
>    - Monthly Cleaning: `"It was documented that all chemical indicators passed, which confirmed the efficient monthly cleaning of all cleanrooms in L-suite."`
>    - Concurrent Positive: `"One other sample (ETX-XXXXXX-XXXX, OOS-XXXXXX) yielded a positive result on [DDMMMYY]. However, that sample was processed by a different processing analyst during a different shift. The two testing events were segregated by shift, processing analyst and primary containment equipment."` (Do NOT include third-party client/product/analyst names; always ask the user if there were any other positives on that day!).
> 6. **GDP 2-Digit Year N/A Sign-Off (`RULE 0Q`)**: Always use 2-digit year `DDMMMYY` (`N/A QYC 21Oct26`), NEVER 4-digit year (`21Oct2026`).
> 7. **Master Table Golden Standards (`RULE 0P`)**: 4 canonical banners without dates, `and` instead of `&`, `No Growth` / `Not Applicable` / `None`, omit Gram staining when Microbial ID exists, empty line between multiple species, uniform `7.0pt` font across all cells & hyperlinks, strict 1-page layout, and 2-file deliverable (`CORP-FORM-21-...pdf` + `Tables OOS-...docx`).








## 🚨🚨 RULE 0A (TOP PRIORITY): Mandatory Analyst Identity Roster, Anti-Hallucination & Ask-First Firewall



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), NEVER guess or hallucinate analyst full names or initials.



1. **The EagleTrax Root Cause (根本病灶)**: EagleTrax logins use `[First Initial of First Name] + [Last Name]` (e.g., `ebennett`, `enioupin`, `mjang`, `dkansara`, `jochoa`, `aodugbesi`). **NEVER expand the first initial by guessing a common first name!** There is **NO Elizabeth Bennett**, **NO Eniola Nioupin**, and **NO Min Jang** in the laboratory!



2. **Official Ground Truth Roster (权威法定花名册 - 必须 100% 严格对应)**:



   - `ebennett` / `elbennett` / `ELB` $\to$ **Ellen Bennett (ELB)** (🚨 严禁写成 Elizabeth Bennett!)



   - `enioupin` / `EN` $\to$ **Elysse Nioupin (EN)** (🚨 严禁写成 Eniola 或 ENI!)



   - `isharma` / `ISS` $\to$ **Ishita Sharma (ISS)** (🚨 官方缩写必须是 ISS，严禁写成 IS!)



   - `acarrillo` / `ANC` $\to$ **Andrew Carrillo (ANC)** (🚨 官方缩写必须是 ANC，严禁写成 AC!)



   - `aodugbesi` / `AOD` $\to$ **Ayomide Odugbesi (AOD)** (🚨 官方全名是 Ayomide Odugbesi，官方缩写是 AOD，严禁写成 Abayomi 或 AO!)



   - `qchen` / `QYC` $\to$ **Qiyue Chen (QYC)** (🚨 官方缩写必须是 QYC，严禁写成 QC 或 KC!)



   - `mjang` / `MJ` $\to$ **Mukyung Jang (MJ)** (🚨 严禁写成 Min Jang!)



   - `gsurber` / `GS` $\to$ **Gabrielle Surber (GS)** (🚨 严禁写成 Gabbie 或混淆为 GA!)



   - `ganyangwe` / `GA` $\to$ **Gerald Anyangwe (GA)** (🚨 与 GS 严格区分!)



   - `dkansara` / `DK` $\to$ **Dhvanir Kansara (DK)** (🚨 严禁写成 David Karangwa!)



   - `jochoa` / `JOC` $\to$ **Jorge Ochoa (JOC)** (🚨 严禁写成 Jose Ocando!)



   - `mbythatagari` / `MRB` $\to$ **Muralidhar Bythatagari (MRB)** (🚨 严禁带中间名 Reddy!)



   - `ksilva` / `KSM` $\to$ **Karla Silva (KSM)** (🚨 严禁写成 Karla Silva Machuca!)



   - `gli` / `GL` $\to$ **Guanchen Li (GL)** (🚨 严禁写成 David Li!)



   - `mnaeem` / `MN` $\to$ **Maryam Naeem (MN)** (🚨 官方确认缩写为 MN!)



   - `ntrinh` / `NT` $\to$ **Nicholas Trinh (NT)** (🚨 官方确认缩写为 NT，Technician 无班次!)



   - `sallen` / `SAA` $\to$ **Sasha Allen (SAA)** (🚨 官方确认缩写为 SAA，严禁写成 SA!)



   - `suprety` / `SU` $\to$ **Sonal Uprety (SU)** | `asaravia` / `ES` $\to$ **Alex Saravia (ES)** (🚨 官方缩写必须是 ES，严禁写成 AS!) | `cdu` / `CCD` $\to$ **Cuong Du (CCD)**



   - `aalanis` / `ALA` $\to$ **America Alanis (ALA)** | `smohammad` / `SMO` $\to$ **Simin Mohammad (SMO)**



   - `ssantamaria` / `SAS` $\to$ **Sophia Santamaria (SAS)** | `restrada` / `RE` $\to$ **Rey Estrada (RE)** | `vsubramanian` / `VV` $\to$ **Varsha Subramanian (VV)**



   - `tduong` / `TD` $\to$ **Thinh Duong (TD)** | `rseymour` / `RS` $\to$ **Robin Seymour (RS)** | `nhulyalkar` / `NH` $\to$ **Neha Hulyalkar (NH)**



   - `cgarza` / `CGS` $\to$ **Clea S. Garza (CGS)** | `dtassew` / `DT` $\to$ **Debrework Tassew (DT)** | `hsmith` / `HS` $\to$ **Halaina Smith (HS)**



   - `tkotisso` / `TK` $\to$ **Tamiru Kotisso (TK)** | `mchukwumerije` / `MC` $\to$ **Maraya Chukwumerije (MC)** | `pgary` / `PG` $\to$ **Pagan Gary (PG)**



   - `karuta` / `KA` $\to$ **Kathleen Aruta (KA)** | `dharrison` / `DH` $\to$ **Domiasha Harrison (DH)** | `dshah` / `DS` $\to$ **Devanshi Shah (DS)**



   - `jesqueda` / `JE` $\to$ **Jacqueline Esqueda (JE)** | `jmehta` / `JM` $\to$ **Jinal Mehta (JM)** | `oajayi` / `OA` $\to$ **Olugbenga Ajayi (OA)**



   - 🚨 **Soniyah Reed 绝对排除**：用户核实其从未在无菌实验室出现过，严禁收录入分析员花名册！



3. **Mandatory Code Firewall**: Always use `resolve_analyst_identity()` and `assert_no_hallucinations()` from `utils.py`.



4. **🚨 严禁自作聪明 · 不懂必问铁律 (Mandatory Ask-First Rule)**: Whenever encountering an unknown EagleTrax username, missing EM logbook, or any uncertain detail, **NEVER guess or extrapolate — STOP immediately and ask the user!**







## 🚨🚨 RULE 0B (TOP PRIORITY): Robin Seymour (RS) Table 2 Section Divider Golden Standard



In all OOS Table 2 documents, the 4 gray section divider banners MUST strictly follow Robin Seymour's verified wording:



1. **`Personnel EM`** (Canonical: NO date suffix! In specific single-day bracketing where QA requested, `Personnel EM for DDMMMYY` is accepted.)



2. **`Biological Safety Cabinet EM for E00XXXX and E00YYYY`** (Canonical: NO `(BSC)`, NO repetitive `BSC`, NO date suffix! In specific single-day bracketing, `(BSC) Biological Safety Cabinet EM for BSC E00XXXX and BSC E00YYYY` is accepted.)



3. **`Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs`** (Must list all room numbers `145, 144, 143, 142` for L-Suite, or `Suite 116 (E001738)`, NO date suffix! Robin Seymour streamlined `(E001979, L-Suite)` to simply `(L-Suite)`.)



4. **`Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs`** (NO date suffix!)







## 🚨🚨 RULE 0C (TOP PRIORITY): Standard Microbiology SOP Designations (MICRO-SOP-2)



In all OOS reports (ScanRDI, Celsis, USP <71>, and EM), environmental monitoring plate incubation and facility monitoring MUST strictly cite **`MICRO-SOP-2`** (`Environmental Monitoring of the Clean-room Facility`), NEVER `SOP 2.600.002`.







## 🚨🚨 RULE 0D (TOP PRIORITY): Zero Guessing on Positive Controls, Reagent Lots & Test Bench Data (阳性对照与核心实验台账零脑补·不明必问铁律)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), NEVER guess, extrapolate, or use dummy template values for Positive Controls or reagent batch data.



1. **cGMP 法定证据链铁律 (Positive Control Legal Ground Truth)**: 阳性对照菌株（Standard / Positive Control Organism，如 `C. albicans`, `S. aureus`, `A. brasiliensis` 等）、批号（Lot #）、效期（Expiration Date）属于实验主管（Robin Seymour）与质保（QA）审核与合规归档的法定关键数据。**严禁从旧案卷拷贝，严禁使用模板默认占位符，严禁擅自脑补！**



2. **强制停步提问门禁 (Mandatory Stop & Ask Gate)**: 凡在 EagleTrax 系统电子记录或现场纸质扫描件中未查明阳性对照菌种及批号效期时，**必须立即停步，向用户主动提问获取**（例如直接向用户询问：“请提供本次检验的阳性对照菌种、Lot 批号与效期”），绝不允许擅自填入虚假或猜测数据导致用户被质保批评！



3. **关键实验参数零脑补延伸**: 包含生物安全柜日常消毒记录批号、双氧水指示卡/IPA 批号、培养基批号效期、设备运行 Session/Record 编号等，凡无法 100% 确认真实性的，一律先向用户确认，宁可多问一句，绝不乱写一字！







## 🚨🚨 RULE 0E (TOP PRIORITY): Strict Prohibition of "&" in Official Report Forms & Dividers (严禁在正式报告表单与标题中使用 "&" 连接符)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), NEVER use the informal ampersand `&` in official forms, equipment lists, section headers, or narrative text.



1. **多套间/多洁净室分行规范**: 在 Cleanroom 认证栏目（如 Page 2 `Text Field32`/`33` 及 Word 表格 `smart_cr_id`），多套间严禁使用 `&` 并列（严禁写成 `E001738 (Suite 116) & E001978 (L-Suite)`）。必须采用标准多行对齐：例如 L-Suite 双房间填报第一行 `E001978 (CR144)`、第二行 `E001979 (CR145)`，右侧对应校准期 `Nov 2026\rNov 2026`（字号 5.0pt）；若涉及 CR116 (`E001738`) 则对应 `Dec 2026`。



2. **正文与标题规范**: 在段落叙述和分段横幅标题中，一律使用标准英文单词 **`and`**（如 `BSC E001312 and E001937`），绝不使用 `&`。杜绝给人以非正式、粗糙或排版怪异的印象。







## 🚨🚨 RULE 0F (TOP PRIORITY): Strict Prohibition of Absolute Claims ("100%") in Narrative (严禁在调查报告中使用 "100%" 绝对化用语)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), NEVER use absolute statements like "100% absence of microbial recovery" or "100% sterile" in narrative text.



1. **cGMP 质保铁律**: 实验主管（Robin Seymour）与质保（QA）审核明确指出此类表述过于绝对。



2. **法定标准用语**: 必须统一规范写作：



   `"Upon review of the environmental monitoring data, absence of microbial recovery (No Growth) was demonstrated across all critical ISO 5 zones and personnel monitoring on the date of testing."`







## 🚨🚨 RULE 0G (TOP PRIORITY): Standard Phrasing for 6-Month Analyte History with Prior Failures & Zero Failures (6 个月特定分析物历史、官方追踪台账溯源与句式铁律)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), when investigating a client's 6-month sample history:



1. **官方受控追踪台账数据源 (Robin Seymour SharePoint Ground Truth)**:



   - 6 个月历史核查必须严格基于官方 SharePoint 受控台账：



     `https://eagleanalytical-my.sharepoint.com/:x:/r/personal/rseymour_eagleanalytical_com/_layouts/15/Doc.aspx?sourcedoc=%7B8A575D4F-67AB-4253-B1E2-08A540577BCB%7D&file=Sterile%20Lab%20-%20OOS%20Tracking%20Log.xlsx&fromShare=true&action=default&mobileredirect=true`



   - **全自动自主穿透检索机制 (Rule 73 终极锁定)**: 依托企业持久化认证会话（%LOCALAPPDATA%\\pastdue_playwright_session），Agent 具备直接通过后台高速 API（UniqueId=8A575D4F-67AB-4253-B1E2-08A540577BCB）秒级并发拉网扫描全部 11 个 Sheet，或通过浏览器自动化（WacFrame_Excel_0）在 Scan RDI OOS、<71> OOS、Celsis Sterility OOS 之间无缝切换、云端回填台账之全套能力。全流程自主完成 6 个月历史检索与句式判定，零猜测零脑补。



   - **问询兜底门禁**: 仅在网络离线或台账连接异常时，才向用户主动提问获取，杜绝擅自断言！



2. **Case A: 无前科句式 (Zero Prior Failures)**:



   - 必须明确包含无菌检验测试方法后缀（如 `under ScanRDI sterility testing.` / `under Celsis sterility testing.` / `under USP <71> sterility testing.`）：



     `"Analyzing a 6-month sample history for [Client Name], this specific analyte "[Analyte Name]" has had no prior Out-of-Specification (OOS) occurrences under [Test Method] sterility testing."`



3. **Case B: 存在前科三步曲句式 (Prior Failures - Robin Seymour 审核金标准)**:



   - **基础陈述句**: `Analyzing a 6-month sample history for [Client Name], this specific analyte "[Analyte Name]" has had [one prior failure / N prior failures], documented under [OOS-ID] for submission [ETX-ID], tested on [DDMMMYY].`



   - **形态学差异对比句**: `The microorganism detected in the previous investigation exhibited [clustered circular-shaped / previous morphology] morphology, which differs from the [current morphology] observed in the current investigation.`



   - **前案调查回顾与闭环评价句**: `The previous failure investigation was reviewed and no laboratory error or assignable laboratory cause was identified.`



4. **语法与标点铁律**: 送检单号 `[ETX-ID]` 之后**必须使用逗号隔开**（`, tested on [DDMMMYY].`），杜绝介词堆叠（`under... for... on...`）。







## 🚨🚨 RULE 0H (TOP PRIORITY): Cleanroom Plural Conjunction Standard (同级多洁净室并列使用 "and" 连词铁律)



**CRITICAL**: When describing weekly active air or surface monitoring across multiple cleanrooms of the same classification:



1. **连词规范**: 必须使用英文连词 **`and`**，严禁使用 `or`（例如：必须写成 `in the ISO 7 cleanrooms (145 and 144)`，绝不可写成 `(145 or 144)`）。



2. **质保依据**: 准确反映所有对应房间均实施了监测且均无微生物生长，杜绝逻辑歧义。







## 🚨🚨 RULE 0I (TOP PRIORITY): ISO 8 Ambient Exposure Technical Defensibility (ISO 8 暴露技术防守用语规范)



**CRITICAL**: Never make absolute statements regarding sample exposure to ISO 8 air (e.g. NEVER write "The test sample does not come into contact with ambient ISO 8 air").



1. **质保依据 (Robin Seymour Review)**: 技术上更具防御性（More technically defensible statement）。



2. **法定标准用语**:



   `"There is minimal exposure of the test sample to ambient ISO 8 air, as samples and supplies are transferred in disinfected, closed containers on carts through the layered cleanroom suites."`







## 🚨🚨 RULE 0J (TOP PRIORITY): Multi-Suite Monthly Cleaning Historical Granularity (多套间月度清洁日期精确拆分法则)



**CRITICAL**: When sample handling or changeover spans across multiple cleanroom suites (e.g. Processing in CR116, Changeover in L-Suite):



1. **严禁混并为单一日期**: 必须精确列出对应套间各自真实的月度清洁消毒日期。



2. **法定标准句式**:



   `"Monthly cleaning and disinfection of the cleanrooms (ISO 7) and their containing biosafety cabinets (ISO 5) were performed on [DDMMMYY] for CR116 and [DDMMMYY] for L-Suite, as per MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. It was documented that all chemical indicators passed, verifying that environmental sanitization remained validated and in control prior to testing."`







## 🚨🚨 RULE 0K (TOP PRIORITY): Equipment Description Room Specificity (设备与洁净室编号精准度规范)



**CRITICAL**: In Page 2 Equipment section (`Text Field32`/`33`) and Word template `smart_cr_id`:



1. **具体房间编号优先**: 避免使用宽泛的 `Suite 116` 或 `L-Suite`，统一使用具体的洁净室代号：



   - L-Suite 双房间标准（Processing in CR145 + Changeover in CR144）：第一行 `E001978 (CR144)`，第二行 `E001979 (CR145)`，对应校准期均为 `Nov 2026\rNov 2026`（5.0pt）

   - 跨套间标准（如 CR116 + CR144）：`E001738 (CR116)`（校准期 `Dec 2026`）与 `E001978 (CR144)`（校准期 `Nov 2026`）







## 🚨🚨 RULE 0L (TOP PRIORITY): Cross-Contamination Opening, Multi-Shift Disclosure & Concluding Standard Phrasing (交叉污染全天批次评估、多班次/多分析员透明披露与结论质保法定用语)



**CRITICAL**:



1. **交叉污染段落标准起始句**:



   统一使用：`"All samples processed on the same day were also evaluated for evidence of sample-to-sample cross-contamination."`



   （严禁使用局限性开篇 `"The microbiological findings were also evaluated..."`，确保评估范围更广更清晰）。



2. **同日其他阳性样本透明披露与隔离证明法则 (Multi-Shift / Multi-Analyst Disclosure Rule)**:



   - 严禁在未核实全天所有班次、所有分析员的情况下草率断言"其他所有人员检验结果均为阴性"。



   - 若当日存在其他阳性样本（即使发生在不同班次、不同分析员、不同套间或不同 BSC）：



     a. **主动透明披露**: 详述该阳性样本案卷号（如 `OOS-XXXXXX`）、送检单号（`ETX-XXXXXX-XXXX`）、客户、药品规格及运行序位。



     b. **极简高效隔离证明法则 (Streamlined Segregation Phrasing - 聚焦不同人员与不同时间)**:
        - 🚨 **用户定稿极简原则**: 无需堆砌过多冗余琐碎信息（如无需专门交代 Prepper 姓名或强调外围套间层级名称）。
        - 核心证明聚焦三大要素：**班次（Shift）、检验处理分析员（Processing Analyst / 样本处理序号）、主操作设备（Primary Containment Equipment / BSC）**。
        - **法定标准精炼句式**:
          `"However, that sample was processed by [Analyst] ([N]th sample processed) during the [Shift]. The two testing events were segregated by shift, processing analyst and primary containment equipment."`



     c. **手套消毒与无菌操作合规重申**: 说明操作人员按 SOP 在样本间严格实施手套消毒与无菌规程。



     d. **排除交叉污染与共用媒介结论**: 证实不存在样本间交叉污染或共同实验室污染源。



3. **结论总结段落法定标准句**:



   统一使用：`"Therefore, the likelihood of laboratory error contributing to this failure is considered minimal, and the original result is deemed valid."`



   （替换刻板句 `"The initial Scan RDI OOS result therefore remains valid in accordance with the applicable laboratory OOS procedure."`）。







## 🚨🚨 RULE 0M (TOP PRIORITY): Official 2-File Deliverable Protocol & Canonical Desktop Destination (最终交付物双文件法定铁律与企业 OneDrive 桌面唯一落盘路径铁律)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), when presenting deliverables to the user upon completing an OOS investigation:



0. **法定桌面落盘路径铁律 (Canonical Desktop Destination Rule)**:
   - 凡是用户或流程要求将 OOS 交付物保存至“桌面（Desktop）”时，**必须严格唯一定位并保存至企业 OneDrive 桌面物理路径**：
     `C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\`
   - 🚨 **严禁仅保存至本地物理桌面** `C:\Users\qchen\Desktop`，彻底杜绝多路径引发的“找不到文件”或版本不同步问题！

1. **用户法定指令 (User Golden Mandate)**:

   用户明确指示，每次交付**严格只向用户输出并提供以下 2 个核心文件**，严禁输出任何多余冗余文件：

   - **文件 1 (带 Table 的完整 PDF 版主案卷 — 官方转录透视命名规范)**:

     - **转录官方表单正式命名**: 凡转录至 ZenQMS 官方空白母表之正式案卷，文件名**必须统一冠以表单编号与同日下载日期前缀**：

       - Phase I: CORP-FORM-21- <DD MMM YYYY> OOS-<ID> <Client> - <Test>.pdf

         （例如：CORP-FORM-21- 01 Oct 2026 OOS-262289 Link Pharmacy (E19112) - ScanRDI.pdf）

       - Phase II: CORP-FORM-22- <DD MMM YYYY> OOS-<ID> <Client> - <Test>.pdf

         （例如：CORP-FORM-22- 10 Sep 2026 OOS-262098 Solyn LLC (E75000) - Phase II.pdf）

     - **核心目的 (一眼核验)**: 使审查人与用户能够**一眼直观核验该案卷是否已原生完成官方表单转录**，以及其采用的母表下载日期是否与阳性发现日期 100% 吻合！

     - 内容包含 6 页 AcroForm 官方调查表单及末尾附加 Page 7/8 Table 1 & Table 2 附表的完整合并版 PDF，严格排版无溢出）。

   - **文件 2 (Word 版独立附表)**: `Tables OOS-<ID> <Client> - <Test>.docx`

     （包含 Table 1 样品信息与 Table 2 环境监测数据的独立 Word 附表文档）。



2. **绝对严禁交付的冗余文件 (Strictly Prohibited Deliverables)**:

   - 🚨 **严禁提供** `Tables OOS-<ID> <Client> - <Test>.pdf`（独立 Table PDF）。

   - 🚨 **严禁提供** `OOS-<ID> <Client> - <Test>.docx`（主表单 Word 版）。

   - 彻底杜绝多余文件造成混淆，保持交付极其精简清晰。



3. **双路径呈现规范 (Dual Path Output Rule)**:

   对上述 2 个文件，必须同时提供：

   - 可点击的 Markdown 超链接（`file:///...`）

   - Windows 原生物理绝对路径（方便用户一键复制粘贴至 Explorer 或 Office 打开）。



---





## 🚨🚨 RULE 0N (TOP PRIORITY): Strict Prohibition of Arbitrary Font Size Alteration & Formatting Tampering during Copy-Paste (严禁在复制粘贴/文本与表格填充时擅自改动任何字号大小与排版格式铁律)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), whenever transcribing, copying, pasting, or inserting user-provided content, table entries, or manual edits into documents/templates:



1. **绝对零擅自改动字号与格式 (Absolute Zero Font & Formatting Tampering)**:

   - 凡用户手动修改、确认或粘贴的内容，**严禁自行更改其字体大小（Font Size）、字体名称（Font Family）、字形（Regular/Italic/Bold）、行间距或对齐方式**！

   - 表格内部各单元格（如 Table 1 样品信息、Table 2 环境监测、Microbial ID、Observation、Notes 等），同类行或同一数据块内的字体大小必须**严格保持完全统一**（例如统一 7.0pt 或 7.5pt），严禁个别行或个别单元格擅自变大变小导致视觉突兀！



2. **物种空行与结构换行严格保留 (Strict Preservation of Empty Lines & Line Breaks)**:

   - 当多菌种（Microbial ID）或多条目按用户要求在物种之间保留空行以增强可读性时，复制粘贴或代码拼接时**必须完整保留原样空行**，绝对严禁擅自压缩、删除空行或合并为单行！



3. **杜绝自作主张的“顺手优化” (No Unsolicited "Formatting Improvements")**:

   - 用户已定稿的排版属于神圣资产（Sacred Ground Truth）。严禁打着“让版面更好看”或“自动适应”的旗号修改字号。若遇特殊空间受限情况，必须**主动停步向用户明确汇报并征得同意**，绝不允许私自偷偷改动！



---





## 🚨🚨 RULE 0P (TOP PRIORITY): Master Table Golden Standards — Robin Seymour (QA) Mandates & User Core Requirements (OOS 附表双轨终极标准规范：Robin 质保法定规范与用户核心排版铁律)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), all Standalone Tables (`Tables OOS-<ID> ... .docx`) and Master PDF Page 7/8 appended tables MUST strictly comply with BOTH Robin Seymour's QA mandates and the User's core requirements:



### 一、Robin Seymour (QA) 质保法定规范 (Robin's QA Golden Standards)

1. **4 大标准分段横幅用语铁律 (Canonical 4-Banner Dividers - 严格无日期后缀)**:

   - Banner 1: `Personnel EM`（官方绝对标准：严禁任何日期后缀！仅在 QA 特殊要求的个别单日回溯案卷中允许带日期）。

   - Banner 2: `Biological Safety Cabinet EM for E00XXXX and E00YYYY`（官方绝对标准：严禁带 `(BSC)` 前缀，严禁重复出现 `BSC`，严禁日期后缀！多台 BSC 之间必须使用 `and`，严禁 `&`）。

   - Banner 3: `Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs`（L-Suite 必须完整列出所有洁燀室编号 `145, 144, 143, 142`，或 `Suite 116 (E001738)`，严禁日期后缀；套间代号统一精简为 `(L-Suite)`）。

   - Banner 4: `Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs`（严格无日期后缀）。

2. **连接词严禁使用 "&" (Strict Prohibition of "&")**:

   - 在横幅标题、设备列表及表内文本中，一律使用标准英文单词 `and`（如 `BSC E001938 and E001937`），绝对严禁使用非正式符号 `&`。

3. **标准观察结果与注释规范 (Observation & Notes Capitalization Standard)**:

   - 阴性结果统一使用 Title Case：`No Growth`（严禁全小写 `no growth` 或首词大写 `No growth`）。

   - 无菌或不适用列统一使用：`Not Applicable`（表格空间允许下严禁简写为 N/A）。

   - 正常无异常注释统一使用：`None`。

   - 异常情况说明首字母大写且末尾不加句号（如 `Plate was desiccated on 5 day read`）。

4. **微生物双名法斜体铁律 (Binomial Nomenclature Italicization)**:

   - 凡涉及属名与种名的微生物学名，必须严格使用斜体（*Italic*），如 *Staphylococcus epidermidis*, *Pseudomonas parafulva*, *Bacillus safensis* 等；属名首字母大写，种名全小写。

5. **附表严格锁定 1 单页 (Strict 1-Page Layout Invariant)**:

   - 独立 Word 附表（Table 1 + Table 2）经排版与 Word COM 导出为 PDF 时，总页数**必须严格锁定为 1 单页**，严禁溢出到第 2 页！合并入主案卷时作为 Page 7 挂载。



### 二、用户核心排版与数据要求 (User's Core Mandates)

1. **有 Microbial ID 则免除染色形态描述 (Microbial ID Precedence - 确证物种免染色形态)**:

   - **“有Microbial ID就不用staining了”**：当送检 EM 平板已获得确证的微生物鉴定结果（Microbial ID 出具具体菌种名）时，Microbial ID 单元格内**直接填报确认的微生物学名，严禁再填充初步革兰氏染色形态**（如无需再写 `Gram (+) cocci` 或 `Gram (-) rods`），避免信息冗余与画蛇添足！

2. **多菌种物种之间强制独立空行 (Mandatory Species Empty Line Separation)**:

   - **“物种要空行”**：当单个 EM 样品检出多种微生物（如 ETX-261012-0433 检出 5 种菌）时，各菌种学名之间**必须强制插入独立空行 (`\n\n`) 进行视觉分隔**，保持层级呼吸感，绝对严禁压缩合并或紧密挨合在单行内！

3. **全表数据单元格字号绝对均一 (Strict Uniform Font Size - 7.0pt / 14 Half-Points)**:

   - **“不过最好字体是都一样大小”**：全表（Table 1 样品信息与 Table 2 环境监测）所有数据单元格——包括 Sampling Site, Frequency, Date, Analyst, Observation, EM Plate ETX Number, Microbial ID, Notes 及超链接文字——其字体大小必须**严格保持绝对均一的 7.0pt (14 half-points)**！

   - 🚨 **底层 XML 超链接防坑铁律**：Word 中的超链接 `<w:hyperlink>` 拥有独立的 `<w:r>` 节点，常规遍历 `p.runs` 会漏改其字号（导致残留 6.5pt / `w:sz w:val="13"`）。必须在底层 XML 级别将所有 `<w:sz>` 与 `<w:szCs>` 统一锁死为 `14`（7.0pt），彻底消除局部字号差异！

4. **原生可点击超链接嵌入 (Native Hyperlink Embedding)**:

   - Table 1 中的主 Sample ID 与 Table 2 中所有检出菌落的 EM 平板 ETX 编号，必须嵌入原生 OpenXML `<w:hyperlink>`，直达 EagleTrax 电子送检详情页。

5. **双文件交付法定归宿 (Rule 0M Compliance)**:

   - 附表最终交付形式严格限定为 Word 版本：`Tables OOS-<ID> <Client> - <Test>.docx`。

   - 严禁单独交付独立的 Table PDF，其导出的 PDF 仅用于作为 Page 7 合并进完整主 PDF 案卷中。





## 🚨🚨 RULE 0Q (TOP PRIORITY): Good Documentation Practices (GDP) N/A Sign-Off Standard (N/A + 分析员官方缩写 + 2位年份日期 DDMMMYY 规范)

**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`) and official AcroForm investigation forms (`CORP-FORM-21`, `CORP-FORM-22`):

1. **适用场景 (Applicability)**:
   - 凡任一文本字段、表格栏目、划线区或溢出续页（如 Page 3 Section C "Other: ________" `Text Field48`、Page 5 "Phase I Summary (continued):" `Text Field51` 等）属于不适用（Not Applicable）并需人工签注划消时，必须严格遵守本条标准。

2. **法定格式 (Canonical GDP Format)**:
   - 必须统一写作：`N/A <Analyst Initials> <DDMMMYY>`
   - 标准范例（以 Qiyue Chen 为例）：`N/A QYC 21Oct26`

3. **2 位年份绝对铁律 (Strict 2-Digit Year Mandate - 严禁 4 位年份)**:
   - 日期中的年份**必须严格使用 2 位数字**（如 `21Oct26`、`01Oct26`、`16Sep26`）。
   - 🚨 **严禁使用 4 位年份**：绝对严禁写成 `21Oct2026`！质保与良好实验室规范（GDP）要求日期格式高度精简统一。

4. **花名册缩写严格对应 (Analyst Roster Invariant)**:
   - 分析员缩写必须 100% 遵守 RULE 0A 官方法定花名册（如 `QYC` 代表 Qiyue Chen，严禁写成 `QC`；`ELB` 代表 Ellen Bennett；`SAA` 代表 Sasha Allen 等）。

5. **格式间距与标点标准 (Spacing & Punctuation)**:
   - 各元素之间使用单一标准英文空格分隔：`N/A` + 空格 + `Initials` + 空格 + `DDMMMYY`。
   - 日期内部无空格无斜杠，末尾绝不添加多余句号或特殊符号。


---





   - **签署日期铁律 (Date of Drafting/Signing Ground Truth)**：N/A 后的签署日期必须严格为**撰写/签署当天的物理真实日期**（Date of Drafting/Signing，例如今天处理即为 22Oct26），**绝对严禁**误填为事故发生日（Date of Incident）或实验检验日（Test Date）！

## 🚨🚨 RULE 0R (TOP PRIORITY): Universal 3-Tier OOS Review Hierarchy & Supervisor (RS) vs QA Role Distinction (全 OOS 模块三级签名流转与主管 RS vs 质保 QA 角色法定铁律)

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
   - Always explicitly list the primary processing BSC (e.g. `ISO 5 BSC E001316`) with its calibration due date (`Jun 2027`).
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


---



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
   - 牢记三级流转：Prepared by: **Qiyue Chen (QYC)** ✍️ $	o$ Reviewed by: **Robin Seymour (RS)** (Micro Lab Supervisor) ✍️ $	o$ Approved by (QA): **Tahsin Kevenk / QA Department** ✍️。

---

## 🚨🚨 RULE 0T (TOP PRIORITY): Sterile Suite & Biosafety Cabinet Analyst Assignment Ground Truth (无菌套间、洁净室与分析员常态化主工位权威映射)

**CRITICAL**: In all sterility investigations (Celsis, ScanRDI, USP <71>), analyst lab locations and BSC assignments strictly follow this official operational ground truth:

1. **Suite 114 Operational Invariant**:
   - **Cleanroom 114B (ISO 7 核心室 / 主操作 BSC E001316)**:
     - Primary Inoculation & Filtration Processing (Processing Analyst).
     - Regular Operating Analysts: **ES (Alex Saravia)**, **GS (Gabrielle Surber)**.
   - **Cleanroom 114A (ISO 7 缓冲室 / 分装换机 BSC E001798)**:
     - 7-Day Incubation Aliquoting & Reading Preparation (Aliquoting Analyst near Celsis Advance 2 E002222).
     - Regular Operating Analysts: **CCD (Cuong Du)**, **ALA (America Alanis)**.

2. **Suite 115 Operational Invariant**:
   - **Cleanroom 115B (ISO 7 核心室 / 主操作 BSC E001309 / E001312)**:
     - High-Throughput Filtration & Sterility Processing.
     - Regular Operating Analysts: **AOD (Ayomide Odugbesi)**, **GL (Guanchen Li)**.

3. **L-Suite (Suite 145 / 144) Operational Invariant (RS Verified Ground Truth)**:
   - **Cleanroom 145 (`CR145 / E001979`, Cert Due `Nov 2026` — ISO 7 核心室 / Scan RDI Processing BSC `E001938`, Cal Due `Jul 2027`)**:
     - Primary Sample Introduction, Pre-labeling & Filtration Processing.
   - **Cleanroom 144 (`CR144 / E001978`, Cert Due `Nov 2026` — ISO 7 缓冲室 / Scan RDI Changeover BSC `E001937`, Cal Due `Jul 2027`)**:
     - Changeover & Sanitization Procedures.
   - Regular Operating Analysts: **ELB (Ellen Bennett)**, **KSM (Karla Silva)**, **VV (Varsha Subramanian)**, **EN (Elysse Nioupin)**.

---


## 🚨🚨 RULE 0V (TOP PRIORITY): Weekly Environmental Monitoring (EM) "Week of Testing" Calendar Alignment Protocol (周度环境监测 "Week of Testing" 自然周严格对齐铁律)

**CRITICAL**: Across all sterility investigations (`Celsis`, `ScanRDI`, `USP <71>`, and `Environmental Monitoring`), NEVER misalign or lag the weekly environmental monitoring dates!

1. **自然周法定同周原则 (Same-Calendar-Week Invariant)**:
   - 实验室周度环境监测（Active Air & Surface Sampling）遵循标准周一至周日（Monday through Sunday）的自然日历周。
   - 附表（Table 2 / Table 3）中的 `Week of Testing`（检验当周），**必须严格采用在该测试日期所在同一个周一至周日自然周内** 所采样的周度监测记录！
   - 例：若样本在周一（05Oct26）进行处理，该周的周度环境监测是在周四（08Oct26）采样，则 `Week of Testing` 必须是 **`08Oct26`**！

2. **严禁整周错位 / 严禁将上一周周检误填为当周 (Strict Anti-Lag / Anti-Shift Rule)**:
   - 严禁将测试日之前一周（T-1 周）的周检记录（如 29Sep26）错填为该测试日的 `Week of Testing`！29Sep26 属于 `Week Before Testing`，绝不能代表 05Oct26 当周！
   - 同理，若在下一周周一（12Oct26）进行分装（Aliquoting），其 `Week of Testing` 必须对应 12Oct26 所在当周（即 15Oct26 周四）的周检记录，严禁拿 08Oct26 当作 12Oct26 的当周！

3. **权威历史案卷证据链严格一致 (Harmonization with Golden Precedents)**:
   - `OOS-262017`: 24Aug26 (Mon) Processing ➡️ Table 2 周检为同周 25Aug26 (Tue)；31Aug26 (Mon) Aliquoting ➡️ Table 3 周检为同周 04Sep26 (Fri)。
   - `OOS-262333`: 05Oct26 (Mon) Processing ➡️ Table 2 周检为同周 08Oct26 (Thu)；12Oct26 (Mon) Aliquoting ➡️ Table 3 周检为同周 15Oct26 (Thu)。

---

## 🚨🚨 RULE 0W (TOP PRIORITY): Global Dual-Track EM Inspection & Growth Focus Screenshot Collaboration Protocol (全局环境监测双轨排查与菌落生长聚焦截图法定协作规范)

**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), NEVER rely solely on automated OCR across raw multi-page scanned PDFs for handwritten corrections or positive plate recoveries.

1. **病灶溯源与核心痛点 (Core Problem Solved)**:
   - 现场纸质环境监测台账（如 `MICRO-FORM-7` / `3.600.002.F06` 浮游菌、`3.600.002.F10` 表面采样）经复印机批量扫描进 EagleTrax 后，常出现方向混杂（横版表格竖向扫描）、低对比度以及重叠手写修正等物理缺陷。
   - 极其关键的微生物法定数据常以手写单线划改更正（如 `0 -> 5 -> 6`）、上标圈码脚注（`①②③④`）及页面边缘小字送检单号（如 `③ ETX-261021-0578, SMO 21Oct26`）呈现。全页压缩 OCR 极易将细小手写笔迹忽略或误判为底色噪点，从而错误推断为“0 CFU / No Growth”，导致严重的质量调查事故。

2. **双轨法定协作机制 (Dual-Track Golden Collaboration Protocol)**:
   - **轨道一：全景后台拉网式调取 (Track 1: Full-Spectrum Background Parsing - Agent 主导)**:
     - 用户将全量原始扫描包上传至 EagleTrax，以确保 cGMP 官方归档与合规证据链完整。
     - Agent 自动登录 EagleTrax 下载全套 PDF，在后台自动提取：
       - 阴性常规数据（全流程人员 Personnel Touch、生物安全柜 BSC 表面与沉降板）；
       - 洁净室月度熏蒸与消毒记录、消毒剂批号效期（HaloMist Lot/Exp）、过氧化氢指示卡批号效期及执行分析员；
       - 洁净套间、生物安全柜、Advance 2 / ScanRDI 设备序列号及工程校准到期日（Dec 2026 / Jun 2027 等）；
       - 培养基（FTM、TSB、Tween 80、TSA 等）批号与效期；
       - 执行人员 Initials 并严格对照 `RULE 0A` 权威花名册 100% 映射官方全名。
   - **轨道二：焦点透视与菌落生长精准截图 (Track 2: Precision Growth Focus & Targeted Screenshot - 用户极简主导)**:
     - 凡在环境监测中遇到任何平皿有菌落生长（CFUs > 0）、复杂划改更正或边缘附注子送检单号（Subculture ETX ID）的页面：
       - **优选方式**：用户随手将含有 Growth 的该单页截一张高清图发送至对话框；
       - **极简方式**：或在对话框中简要发送一句话说明（例如：*“P23 ISO 8 浮游菌有 6 CFU，单号 ETX-261021-0578”*）。
     - Agent 接收后，立即将该截图/说明作为最高法定真实锚点（Ground Truth Anchor），精确核实：
       - 最终确证 CFU 数值（如初读 5 终读 6 CFU）；
       - 48 小时与 5 天读板人签名与真实读板日期时间；
       - 边缘登记的送检单号（如 `ETX-261021-0578`）；
       - 复核人签核（Reviewer By/Date）。

3. **强制闭环防护与质保防御论证 (QA Defense Integration)**:
   - Agent 依据用户提供的 Growth 截图，自动在附表 Table 2 / Table 3 中精准填报（如 Observation `6 CFUs`，EM Plate ETX Number `ETX-261021-0578`，Microbial ID `Pending`，Notes `None`）。
   - 在调查报告正文（Page 5）中，自动应用 `RULE 0I` 等质保防守金标准，详述该阳性平皿严格局限于最外层 ISO 8 更衣缓冲区，阐明套间级联外向正压差、物理气闸密闭门及消毒转运容器之有效隔离，强力论证核心 ISO 7 洁净室与 ISO 5 操作区绝对 100% 阴性受控，排除实验室污染。

4. **强制停步提问门禁 (Mandatory Stop & Ask Gate)**:
   - 即使有截图，若局部手写字迹模糊、缩写存疑（例如无法确定是 MC 还是 MN）或单号残缺，Agent 必须恪守 `RULE 0A` 与 `RULE 0D`，立即圈出疑问区域向用户提问求证，绝对严禁任何擅自推测或脑补！



## 🚨🚨 RULE 0X (TOP PRIORITY): Robin Seymour (RS) Quality Review & Defensibility Standards (ScanRDI 案卷与套间布局终极质保规范)

**CRITICAL**: Across all ScanRDI and cleanroom OOS investigations, all generated forms and narratives MUST strictly incorporate Robin Seymour's verified Quality Standards:

0. **Robin Seymour (RS) 主管地位与执行铁律 (Supervisor Authority Mandate)**:
   - **组织关系**: **Robin Seymour (RS)** 是报告撰写人 **Qiyue Chen (QYC)** 的直接实验室主管（Microbiology Laboratory Supervisor）。
   - **最高执行准则**: RS 在实验室日常管理、技术规范、套间物理布局、理化证据链审查及官方审核意见上具备最高权威与法定效力。在所有 OOS 调查案卷中，我们**必须 100% 严格遵从 RS 的要求与审核修改**，严禁擅自变动，严禁使用旧模板覆盖 RS 确立的新标准！

1. **原始数据包事件数唯一法定原则 (Event Count Source Truth)**:
   - 严禁仅采信分析员手写便签或口头流转数据（Analyst data entries）。**必须严格调取 Scan Data Packet 仪器原始打印报告核验总事件数（Total Events）与确证微生物事件数（Confirmed Microbial Events）**。
   - 例如 OOS-262289 中分析员记录为 15，实际原始数据包显示为 19 total events, 4 confirmed microbial events，以原始数据包 19 为唯一法定依据。Table 1 与正文叙述必须绝对严格一致。
   - 🚨 **用户协作与主动提问门禁 (User Collaboration Gate)**: Scan Data Packet 原始打印件在系统或扫描附件中往往难以自动定位。Agent **绝不可自行盲目采信手写便签或胡乱推测，必须主动向用户求助提问**（如：*“Scan Data Packet 原始数据包较难定位，请帮我核验本次检验的真实 Total Events 与 Confirmed Microbial Events”*），由用户协助查找并提供，确保数据 100% 真实准确！

2. **L-Suite 物理房间真实对应布局 (L-Suite Room Mapping)**:
   - **BSC E001938** 位于 **Cleanroom 145 (CR145)**（用于样品引入、预贴签与过滤处理）。
   - **BSC E001937** 位于 **Cleanroom 144 (CR144)**（用于换批消毒清场 Changeover）。
   - 🚨 **严禁混为一谈**：绝不可将两台 BSC 笼统写在 CR144！必须精准体现各 BSC 所在的真实房间编号。

3. **双套间认证与校准期精准填报 (Environmental Certification Precision)**:
   - Cleanroom 认证（Page 2 `Text Field32`）：必须同时填报 `E001978 (CR144)\rE001979 (CR145)`（字号 5.0pt）。
   - 校准有效期（`Text Field33`）：对应 `Nov 2026\rNov 2026`（字号 5.0pt）。
   - BSC 校准有效期（`Text Field44`）：E001938 与 E001937 均填报 `Jul 2027`。
   - ScanRDI 仪器校准期（`Text Field31`/`35`）：填报 `Oct 2026`。

4. **微生物形态学对比科学严密性铁律 (Morphology Defensibility Standard)**:
   - 当环境菌（如 *Staphylococcus epidermidis*）同属球菌形态时，**严禁宣称其形态与阳性样本 Cocci 截然不同**！
   - 必须完全删除形态差异对比，将防守论点纯粹建立在“物理套间硬隔离、闭门、正压梯度气流屏障，以及 ISO 5 台面和手套完全无菌（No Growth）”之上。

5. **月度清洁与交叉污染极简精炼句式 (Streamlined Quality Phrasing)**:
   - 月度清洁：`"It was documented that all chemical indicators passed, which confirmed the efficient monthly cleaning of all cleanrooms in L-Suite."`
   - 同日阳性：`"One other sample (ETX-XXXXXX-XXXX, OOS-XXXXXX) yielded a positive result on [Date]. However, that sample was processed by a different processing analyst during a different shift. The two testing events were segregated by shift, processing analyst and primary containment equipment."`

---



## 🚨🚨 RULE 0Y (TOP PRIORITY): Official L-Suite Cleanroom & Equipment Layout Ground Truth (Figure 6 - H2O2 Strips & Equipment Mapping)

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







### 16. Robin Seymour (Microbiology Supervisor) Sterility OOS Review Standards & Harmonization Rules (Robin 官方复核标准与合规八项铁律)



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







### 26. Analyst Legal Full Name Precision & Universal Tri-Module Ground Truth Standard (全部 3 种 OOS 模块分析员法定全名与官方缩写守恒铁律)



**CRITICAL**: Across ALL 3 sterility OOS modules (`ScanRDI`, `Celsis`, `USP <71>`) and EM investigations:



1. **Analyst Legal Name Accuracy (Zero Tolerance for Hallucination & Initials Precision)**:



   - For EagleTrax username `enioupin`: The legal system name is **Elysse Nioupin (EN)** (🚨 **ABSOLUTE RED LINE**: NEVER hallucinate "Eniola Nioupin" or invent "ENI"!).



   - For EagleTrax username `isharma`: The legal initials are strictly **ISS** (**Ishita Sharma (ISS)**, NEVER "IS"!).



   - For EagleTrax username `acarrillo`: The legal initials are strictly **ANC** (**Andrew Carrillo (ANC)**, NEVER "AC"!).



   - For EagleTrax username `aodugbesi`: The legal name is strictly **Ayomide Odugbesi (AOD)** (🚨 **ABSOLUTE RED LINE**: NEVER hallucinate "Abayomi Odugbesi" or invent "AO"!).



   - For EagleTrax username `qchen`: The legal initials are strictly **QYC** (**Qiyue Chen (QYC)**, NEVER "QC" or "KC"!).



   - For initials `MJ` in sample prep: The full legal system name is **Mukyung Jang (MJ)** (NEVER shorthand or truncated "Min Jang").



   - For initials `GS` / `gsurber`: The full legal name is **Gabrielle Surber (GS)** (NEVER "Gabbie Surber", NEVER confused with "GA").



   - For initials `GA` / `ganyangwe`: The full legal name is **Gerald Anyangwe (GA)** (NEVER confused with Gabrielle Surber).



   - For initials `DK` / `dkansara`: The full legal name is **Dhvanir Kansara (DK)** (🚨 NEVER hallucinate "David Karangwa"!).



   - For initials `JOC` / `jochoa`: The full legal name is **Jorge Ochoa (JOC)** (🚨 NEVER hallucinate "Jose Ocando"!).



   - For initials `INS` / `isunesara`: The full legal name is **Inaara Sunesara (INS)**.



   - For initials `MRB` / `mbythatagari`: The full legal name is **Muralidhar Reddy Bythatagari (MRB)**.



   - Global Concordance: Analyst full names and initials must be strictly harmonized across Section B comments (Text Field13), Page 1 personnel box (Text Field3), Page 3-5 narratives (Text Field49/50/51), and master Word/PDF report tables.



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







### 33. EagleTrax Benchmark Extraction & Raw Bench Operator Ground Truth Rule (EagleTrax 原始凭证穿透与实验台账真实操作人校准规范 - 🚨 永久铁律)



**CRITICAL**: In sterility OOS investigations (such as `OOS-262098` / `ETX-260902-0505`):



1. **Raw Bench Operator Penetration (前置台账真实人员穿透提取)**:



   - Always extract exact bench operators and timestamps from EagleTrax `#EventHistory` and logbooks.



   - For `ETX-260902-0505`: Prepping Analyst is `isharma` (**Ishita Sharma, ISS**) on 03Sep26 20:48; Processing Analyst is `asaravia` (**Alex Saravia, ES**) on 04Sep26 22:33; Reading Analysts are `enioupin` (**Elysse Nioupin, EN**) and `acarrillo` (**Andrew Carrillo, AC**).



   - In Form 1 (`CORP-FORM-21`), Section A, narratives, and interviews, always specify:



     - `Prepping Analyst: Ishita Sharma (ISS)`



     - `Processing Analyst: Alex Saravia (ES)`



     - `Reading Analyst: Elysse Nioupin (EN)`



2. **Sample Filtration Quantity Calibration (样品过滤量真实规格校准)**:



   - Extract exact filtration volume from Worksheets: `45 gm added per media` across `3 vials` (each container 30 mL/gm, 90 gm total batch, 45 gm filtered into FTM, 45 gm filtered into TSB). State this explicitly in narrative paragraph 8.



3. **Consumables Traceability (耗材系统批号无缝闭环)**:



   - Accurately map FTM Lot `2567320` (Exp `12/16/2026`), TSB Lot `2548010` (Exp `04/19/2027`), and Fluid D Lot `685561` (Exp `10/31/2026`).







### 34. Mandatory Standard Analyst Name & Initials Roster & Reverse Mapping Architecture Firewall (全项目统一分析员全名与缩写权威词典与技术防火墙 - 🚨 终极铁律)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), NEVER guess or hallucinate analyst full names or initials.



- **Root Cause of Past Hallucinations (根本病灶)**: EagleTrax system logins follow the format `[First Initial of First Name] + [Last Name]` (e.g., `ebennett`, `enioupin`, `mjang`, `dkansara`, `jochoa`, `aodugbesi`). Previously, agents saw the single initial `e-` or `m-` or `a-` and hallucinated common first names (`Elizabeth`, `Eniola`, `Min`, `David`, `Jose`, `Abayomi`). **There is NO Elizabeth Bennett, NO Eniola Nioupin, NO Min Jang, NO Abayomi Odugbesi in the laboratory!**



- **Absolute Rule**: NEVER expand an EagleTrax username by guessing the first name! Always query `utils.py` (`resolve_analyst_identity()`). If an unknown username appears, STOP and ask the user.



1. **Critical High-Risk Corrections (高危易错人员绝对锁定)**:



   - `ELB` / `ebennett` / `elbennett` -> **Ellen Bennett (ELB)**（🚨 **EagleTrax 账号是 ebennett，严禁脑补为 "Elizabeth Bennett"！官方全名必须是 Ellen Bennett！**）



   - `EN` / `enioupin` -> **Elysse Nioupin (EN)**（🚨 **严禁写成 "Eniola"！官方全名必须是 Elysse Nioupin！**）



   - `ISS` / `isharma` -> **Ishita Sharma (ISS)**（🚨 **官方缩写必须是 ISS！**）



   - `ANC` / `acarrillo` -> **Andrew Carrillo (ANC)**（🚨 **官方缩写必须是 ANC！**）



   - `AOD` / `aodugbesi` -> **Ayomide Odugbesi (AOD)**（🚨 **官方全名必须是 Ayomide Odugbesi，官方缩写必须是 AOD，严禁写成 Abayomi！**）



   - `ES` / `asaravia` -> **Alex Saravia (ES)**（🚨 **官方缩写必须是 ES，严禁写成 AS！**）



   - `QYC` / `qchen` -> **Qiyue Chen (QYC)**（🚨 **官方缩写必须是 QYC！**）



   - `MJ` / `mjang` -> **Mukyung Jang (MJ)**（🚨 **严禁写成 "Min Jang"！官方全名必须是 Mukyung Jang！**）



   - `GS` / `gsurber` -> **Gabrielle Surber (GS)**（🚨 **严禁混淆为 "GA" 或 "Gabbie"！官方缩写必须是 GS！**）



   - `GA` / `ganyangwe` -> **Gerald Anyangwe (GA)**（🚨 **官方缩写必须是 GA，与 GS 严格隔离！**）



   - `DK` / `dkansara` -> **Dhvanir Kansara (DK)**（🚨 **严禁脑补为 "David Karangwa"！官方全名必须是 Dhvanir Kansara！**）



   - `JOC` / `jochoa` -> **Jorge Ochoa (JOC)**（🚨 **严禁脑补为 "Jose Ocando"！官方全名必须是 Jorge Ochoa！**）



   - `INS` / `isunesara` -> **Inaara Sunesara (INS)**（🚨 **官方缩写必须是 INS！**）



2. **Codebase Single Source of Truth & Interceptor Firewall (底层代码强制拦截防火墙)**:



   - `utils.py` 中的 `clean_analyst_name()` 和 `resolve_analyst_identity()` 会自动将 `elizabeth bennett` / `elizabeth` 强制拦截纠正为 `Ellen Bennett`。



   - `utils.py` 中的 `assert_no_hallucinations()` 门禁会在文件输出前全面扫描，一旦检测到任何历史臆想人名立即抛出致命异常阻断程序。







3. **🚨 绝对严禁自作聪明 · 不懂必问铁律 (Mandatory Ask-First Rule on Any Ambiguity / Unknowns - 终极原则)**:



   - **零脑补 (Zero Guessing)**: 严禁自作聪明！无论是在分析员姓名、缩写、耗材批号、设备编号、培养箱校准日期、还是洁净室房间拓扑上，如果遇到任何不确定、模糊、或者在权威字典和台账中找不到对应关系的信息，**绝对严禁自行臆测填写**！



   - **立即停步请示 (Stop & Ask Immediately)**: 只要遇到花名册未收录的 EagleTrax 账号、未见记录的台账、或不确定的调查细节，必须**立刻停下手中工作，在对话框中直接向用户提问请示**！



   - **宁可多问一句，绝不乱写一字**: 在 cGMP 合规审计与 OOS 调查中，多问一句是专业严谨，乱猜一字是致命事故。用户的指示永远是唯一依据，严禁任何自作聪明！







### 35. Robin Seymour (RS) Official OOS Table Divider Golden Standard (Table 2 灰色分段横幅终极标杆 - 🚨 终极锁定)



**CRITICAL**: In all microbiology OOS investigations (such as `OOS-262194` and `OOS-262150`), Table 2 section divider rows (gray banners) MUST strictly adhere to Microbiology Supervisor Robin Seymour's official verified syntax:



1. **Personnel EM Divider**: Strictly **`Personnel EM`** (Canonical: NO date suffix, NEVER write `Personnel EM for DDMMMYY` unless specific single-day QA mark is requested).



2. **Biological Safety Cabinet Divider**: Strictly **`Biological Safety Cabinet EM for E00XXXX and E00YYYY`** (Canonical: NO `(BSC)`, NO repetitive `BSC`, NO date suffix).



3. **Weekly Active Air Divider**:



   - For L-Suite: Strictly **`Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (E001979, L-Suite) with Processing BSCs`** (or `with Changeover BSC` / `with Processing and Changeover BSCs`). Must explicitly list all cleanroom rooms (`145, 144, 143, 142`), NO date suffix.



   - For Cleanroom Suite 116: Strictly **`Weekly Active Air Sampling of Cleanroom Suite 116 (E001738) with Processing BSC`** (NO date suffix).



4. **Weekly Surface Divider**:



   - Strictly mirror the cleanroom room list from the active air section, e.g. **`Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (E001979, L-Suite) with Processing BSCs`** (NO date suffix).







### 36. Robin Seymour (RS) Golden Benchmark Harmonization Rules (Robin 质保复核黄金标杆与全局通用八项铁律 - 🚨 终极锁定)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), all generated investigation packages MUST strictly harmonize with Microbiology Supervisor Robin Seymour's signed benchmark standards (`OOS-262150 BelieveRX Pharmacy Signed by QYC`):







1. **Page 4 Combined Processing & Changeover Phrasing (处理与换台同人员合并句式)**:



   - When the same analyst performs both processing and changeover, DO NOT write repetitive separate clauses.



   - Use Robin's combined clause: `"processing and change over steps were performed by [Analyst] ([Nth] sample processed)"` (e.g., `"processing and change over steps were performed by Ellen Bennett (4th sample processed)"`).







2. **Page 3 Analyst Name Hyphen Wrapping in Container Integrity Paragraph (容器完整性段落分析员连字符规范)**:



   - In narrative paragraph 2 (container integrity and physical inspection), analyst names must be wrapped in spaced hyphens:



     `"Prepping analyst - [Name] - and processing analyst - [Name] - verified the physical integrity and packaging of the sample..."`



     (e.g., `"Prepping analyst - Mukyung Jang - and processing analyst - Ellen Bennett - verified..."`). NEVER write bare unhyphenated names.







3. **Page 4 EM Evaluation Block Hierarchy & Native Returns (`Text Field50` 环境监测三层区块排版与单 `\r` 子标题规范)**:



   - In AcroForm multiline field `Text Field50`, split EM review narrative into 3 distinct `\r \r`-delimited blocks with single `\r` between sub-header and body text:



     - Block 1 (Standalone Header): `"Environmental Monitoring (EM) Recoveries Review:"`



     - Block 2 (ISO 5 BSC & Personnel): `"ISO 5 BSC and Personnel Evaluation:\r[body text]"`



     - Block 3 (Cleanroom EM): `"Cleanroom EM Evaluation:\r[body text]"`



   - Using a single `\r` between sub-header and body ensures sub-headers stay tightly coupled to their paragraph and prevents Adobe Acrobat double-newline rendering issues.







4. **Page 2 Equipment Block Must Include BSCs (`Text Field43` & `Text Field44` 生物安全柜全量录入规范)**:



   - In addition to incubators and sensors, `Text Field43` (Equipment ID) and `Text Field44` (Calibration Due Date) MUST ALWAYS include both the processing BSC and changeover BSC (e.g., `BSC E001938` / `BSC E001937`) along with their calibration due dates (`Feb 2027` / `Feb 2027`).



   - Maintain 1-to-1 horizontal blank line alignment (`\r \r`) between equipment IDs and calibration due dates.







5. **Page 2 Material Lot/Exp Field Formatting (`Text Field22` 原生回车换行规范)**:



   - Reagent lot and expiration blocks must use native Acrobat line breaks (`\r` within each group and `\r \r` between different reagent/media groups), avoiding raw `\n` to prevent font shrinkage and baseline clipping.







6. **Page 6 Prepared By Print Field (`Text Field53` 调查员纯粹性规范)**:



   - `Text Field53` (Prepared by Print) must strictly be the report author / investigator's name alone (`"Qiyue Chen"`). NEVER concatenate or prefix with bench operators (NEVER write `"Ellen Bennett (Written by: Qiyue Chen)"`).







7. **Table 2 Weekly EM Day/Week Column Standard (周检周期表头规范)**:



   - Column `Day /Week(s)` must strictly be `"Week of Testing"`. NEVER append `"\n(Scan)"` or `"Week on Testing Date"`.







8. **Client 6-Month Prior Failure Count Audit (客户先前真实阳性次数审计)**:



   - In client 6-month historical reviews, only count genuine confirmed prior sterility test failures/OOS events. Any cancelled, voided, or invalidated historical records MUST NOT be counted in the numerator.







### 37. Sterility Investigation Architecture: Only-Phase-I vs. Phase-II Master Tables Distinction (仅有 Phase I 与包含 Phase II 案卷表格根本性区分规范 - 🚨 终极锁定)



1. **为什么有的案子表格 Table 1 只有 1 个样本？**



   - 因为那些案子（如常规日常 OOS、调查判定为实验室错误直接结案、或客户未要求复测的案子）**只有 Phase I**，物理上根本没有发生复测（No Retest / No Phase II）！所以 Table 1 自然只有 1 个样本！



2. **为什么包含 Phase II 的案子 Table 1 必须呈现 2 个样本？**



   - 凡是一起无菌检验 OOS 开启并存在 Phase II（即发生了 Retest 复测，如 `OOS-262098` 与 `OOS-261987`），调查母表（Master Tables）**必须始终统筹呈现两阶段的两个测试样本**（Row 1: 初始测试样本；Row 2: 复测测试样本）！



   - 无论是在 Phase I 报告后附带的表格，还是 Phase II 报告后附带的表格，表格附件均为统一的 Master Tables 完整母表（Table 1 双样本 + Table 2 初始环境 + Table 3 复测环境）！



3. **严禁混淆与机械类比：**



   - 绝不能看到“其他只有 Phase I 的案子表格只有 1 个样本”，就盲目将“有 Phase II 的案子”的 Phase I 附件表格强行裁剪成单样本！两类案子的业务场景截然不同！对于存在 Phase II 的 OOS 案卷，Table 1 必须全量统筹双样本呈现。







### 38. Native Auto Font Size (fs=0) for ALL PDF Form Transcriptions (全 OOS 模块 PDF 表单转录原生 Auto 字号永久铁律 - 🚨 终极锁定)



**CRITICAL**: Across all OOS modules (`ScanRDI`, `Celsis`, `USP <71>`, and `Environmental Monitoring`), when transcribing data into official PDF forms (such as `CORP-FORM-21 / 3.100.019.F01` and `CORP-FORM-22 / 3.100.019.F02`):



1. **Root Cause of "Locked Font Size" Bug (病灶与血的教训)**:



   - Scripts using PyMuPDF (`fitz`) that manually set `w.text_fontsize = <fixed_value>` overwrite the PDF's native `/DA` appearance string from `Helv 0.000 Tf` (Auto) to a fixed font size (e.g. `Helv 8.5 Tf`).



   - This locks field sizes, preventing Foxit PhantomPDF / Adobe Acrobat from dynamically autoscaling text when line breaks or edits occur, destroying the authentic human copy-paste experience.



2. **Strict Ban on Hardcoded Font Sizes via fitz (严禁在 fitz 中写死文本字号)**:



   - NEVER set `w.text_fontsize = <positive_number>` on text fields.



   - Text fields must be filled via `pypdf`'s `update_page_form_field_values()`.



   - `fitz` widget updates are ONLY allowed for checkboxes (toggling `/Yes` / `Off`).



   - If a text widget must be touched via fitz in extreme circumstances, explicitly force `w.text_fontsize = 0` (Auto).



3. **Mandatory NeedAppearances Flag (强制开启外观动态重构)**:



   - Always call `writer.set_need_appearances_writer(True)` on `pypdf.PdfWriter` to instruct PDF viewers to dynamically render appearance streams on open.



4. **Official Blank Form Native Baseline (官方原生 Auto 基准)**:



   - Blank ZenQMS official forms natively have `fs=0.0` (Auto) across all text fields. Transcriptions MUST preserve this native state 100%.







### 39. Phase II Retest Investigation & Exact Phase I Verbatim Transcription Standard (Phase II 调查报告与 Phase 1 叙述逐字逐句 100% 镜像转录铁律 - 🚨 终极锁定)



**CRITICAL**: When preparing or transcribing Phase II investigations (`CORP-FORM-22 / 3.100.019.F02`):



1. **Verbatim Phase I Transcription (Phase 1 叙述 100% 逐字原貌复制)**:



   - Under the header `INITIAL TEST UNDER SUBMISSION [ETX-XXXXXX-XXXX]:`, the entire narrative text MUST be copied character-by-character from the finalized Phase I report (`CORP-FORM-21`).



   - Retain all exact line breaks, empty paragraph lines, punctuation, and wording. NEVER summarize, condense, or alter Phase I text.



2. **Strict Demarcation of Retest Narrative (复测部分严格以独立标题起首)**:



   - The new investigation for the retest begins strictly under the header: `PHASE II RETEST UNDER SUBMISSION [ETX-XXXXXX-XXXX]:`.



3. **Empty Line Visual Ergonomics (非连续段落空行美学)**:



   - Between distinct components (headers, testing phase summaries, structured lists), always insert an empty line (`\r\n\r\n`) so that text renders cleanly and beautifully in auto-resizing PDF form boxes.







### 40. Elimination of Robotic AI All-Caps Headers & Over-Explaining Conciseness Rule (严禁机械全大写 AI 标题与“写得越多越错”精简防御铁律 - 🚨 终极锁定)



**CRITICAL**:



1. **No Robotic All-Caps Headers (严禁机械全大写小标题)**:



   - NEVER insert artificial all-caps titles like `EVALUATION OF RETEST FINDINGS AND JUSTIFICATION OF PRIOR SCANRDI RESULT:`. Use natural, organic transition sentences (e.g. `In evaluating the prior rapid ScanRDI testing for Lot ...`).



2. **Conciseness Over Speculative Essays ("写得越多越错"原则)**:



   - When evaluating why a prior rapid ScanRDI test passed while USP <71> compendial testing failed, DO NOT write speculative essays about sample volume, lack of method suitability, or statistical vial probability.



   - Use the concise, defensible laboratory conclusion:



     `"...attributed to non-uniform microbial distribution at low concentrations within the compounded batch."`



   - More speculative text introduces audit liabilities. Keep it concise, focused, and defensible.







### 41. Preservation of Native Digital Signature Fields in PDF Forms (CORP-FORM-21 / CORP-FORM-22 原生数字签名域绝对禁止转为普通文本铁律 - 🚨 终极锁定)



**CRITICAL**:



1. **AcroForm Architecture Reality**: In `CORP-FORM-22` (and `CORP-FORM-21`), Section F (`Section F – Acknowledgement and Closure`) fields under `Prepared by:`, `Reviewed by:`, and `Approved by (QA):` are native Digital Signature widgets (`/FT /Sig`), namely `Signature_0`, `Signature_1`, and `Signature_2`.



2. **Strict Ban on Field Type Conversion (严禁将 Signature 篡改为 Text)**:



   - NEVER convert `Signature_0`, `Signature_1`, or `Signature_2` to `/FT /Tx` (Text widgets).



   - NEVER inject plain string values into signature fields. Overwriting them strips Foxit's native digital certificate signature interaction (purple box and red/orange "SIGN" flag).



3. **Execution Protocol (标准执行范式)**:



   - Keep `Signature_0` as a native, blank Signature widget (`/FT /Sig`). The analyst will digitally sign via Foxit with their personal digital certificate, automatically generating "Digitally signed by Qiyue Chen" and certificate metadata.



   - Populate `Date Field3` with the investigation date (`16-Oct-2026`).



   - `Reviewed by:` and `Approved by (QA):` (`Signature_1`, `Signature_2`, `Date Field4`, `Date Field5`) must remain completely blank.







### 42. Phase II Section E Scope & Permission Protocol (Phase II Section E 调查范围与权限门禁规范 - 🚨 终极锁定)



**CRITICAL**:



1. **Role Demarcation**: In `CORP-FORM-22`, `Section E – Corrective Actions and Preventive Actions / Additional Evaluation` is reserved for QA management / supervisor disposition.



2. **Default Leave Blank**: Analysts/investigators must leave Section E empty/unfilled by default, unless explicit CAPA instructions are provided in writing by QA/management.







### 43. Proactive Official ZenQMS Form Transcription Reminder (主动提醒官方表单转录常态化铁律 - 🚨 终极锁定)



**CRITICAL**: Every new OOS investigation requires filling out a freshly downloaded, same-day ZenQMS official PDF form (`CORP-FORM-21 / 3.100.019.F01` for Phase I, or `CORP-FORM-22 / 3.100.019.F02` for Phase II). Manual entry into 150+ fields is tedious and error-prone.



1. **Mandatory Proactive Reminder (完成即提醒门禁)**:



   - Whenever drafting, data reconciliation, or Word report generation is completed, the agent **MUST ALWAYS proactively remind the user to perform the official PDF transcription**.



   - Standard Reminder:



     `"💡【官方表单转录提醒】：检测叙述与台账数据已全部就绪！请问你今天从 ZenQMS 下载了最新空白官方 PDF 表单（CORP-FORM-21 或 CORP-FORM-22）吗？如果有，请直接将文件或路径告诉我，我将立即为你自动执行 100% 官方表单转录、自动字号匹配（Auto Font Size）与附件表格合成！"`



2. **Auto-Detection**: If a freshly downloaded blank form already exists in Desktop, Downloads, or repository directories, the agent should proactively offer to execute the transcription script immediately.

### 44. Robin Seymour (Microbiology Supervisor) Official Review Golden Standards for USP <71> Phase I & Phase II (Solyn LLC OOS-262098 质保复核定稿标杆与六大黄金准则 - 终极锁定)
**CRITICAL**: In all USP <71> OOS investigations (such as `OOS-262098 Solyn LLC`), reports MUST strictly incorporate Microbiology Supervisor Robin Seymour's verified edits:
1. **Equipment Section "Phase I Summary" & Mandatory ISO 5 BSC (设备栏引用与生物安全柜强制罗列)**:
   - In `CORP-FORM-21` Equipment section, when referring to the investigation summary for environmental controls, use `"Phase I Summary"` (NEVER write `"See Phase I Summary"`).
   - In Equipment table, always explicitly list the primary processing BSC (`ISO 5 BSC E001316`) with its calibration due date (`Dec 2026`).
2. **Mandatory Reconstitution Process Narrative (冻干粉/小容量注射剂复溶重构过程详述铁律)**:
   - In Phase I narrative paragraph 2/3 (`Text Field49`), when testing lyophilized vials or powders, **MUST explicitly detail the reconstitution process**:
     `"...reconstituted with 2mL Bacteriostatic Water per vial (each vial containing GLP3R/Cagrilinitide in 25mg/5mg concentration). Half the contents of each reconstituted vial was then dispensed directly from the container into sterile canisters placed on a vacuum manifold for membrane filtration."`
     (NEVER vaguely summarize as `"45 gm filtered per media across 3 vials"`).
3. **Cleanroom Suite 115 Certification Due Date Standard (CR115 校准期统一标准)**:
   - Cleanroom Suite 115 (`CR 115 / E001737`) certification due date is strictly **`Dec 2026`** (identical to Suite 114 / `Dec 2026`), NEVER June 2027.
4. **RS Golden 2-Paragraph Defense for Prior Passing Rapid Test (先前快速法合格与当前药典法不合格冲突时的质保官方两段式防守金标准)**:
   - When a client previously submitted the same lot for rapid testing (e.g. ScanRDI) with passing results, but USP <71> compendial initial and retest fail:
     - **Paragraph 1 (Method Suitability Defense - 方法适用性缺失防守)**:
       `"It was identified that the same sample lot (2608-216) had previously been submitted by the client for ScanRDI sterility testing under submission ETX-260807-0602 and yielded passing results. However, it is important to note that a review of the submission determined that no approved method suitability test was available to support the validity for microbial growth recovery using the ScanRDI method for this specific sample."`
     - **Paragraph 2 (Non-Homogeneous Bioburden Distribution Defense - 低浓度微生物非均一分布防守与不可比定论)**:
       `"Furthermore, the discrepancy between the previous passing ScanRDI result and the current USP <71> sterility failure may be attributable to the non-homogeneous distribution of microbial contamination within the compounded batch. When microbial contamination is present at very low levels, microorganisms may not be uniformly distributed throughout the product, resulting in variability between samples collected from the same lot. Consequently, one aliquot may not contain contaminating microorganisms and yield an acceptable result, while another aliquot from the same batch may contain viable microorganisms and generate a positive sterility test result. Given the absence of a qualified method suitability study for the ScanRDI submission, the prior passing result cannot be considered directly comparable to the current USP <71> sterility test outcome."`
5. **Phase II Root Cause Statement Simplification (Phase II 根因分析极简法定用语)**:
   - In `CORP-FORM-22` Section D (`Text Field25`), NEVER include lengthy recovery kinetics or ScanRDI discussions. Focus strictly on the confirmed recovery of the identical organism during retest:
     `"Inherent sample contamination within Lot 2608-216, supported by recovery of the same organism (Microbacterium sp. PM5) during retesting. The original USP <71> sterility failure was confirmed as valid."`
6. **Investigation Final Conclusion Wording (调查结语权威法定表述)**:
   - Conclude by confirming inherent product contamination and explicitly stating that **both initial and retest failures are considered valid**:
     `"Therefore, inherent product contamination in Lot 2608-216 is confirmed, both the initial failing result for ETX-260902-0505 and the retest ETX-260914-0470 is considered valid."`
7. **Section E Supervisor Disposition (Section E 主管与质保处置权定论)**:
   - `CORP-FORM-22` Section E is the sole jurisdiction of QA/Supervisor. Analysts must leave Section E empty during drafting. Robin Seymour directly stamped `"N/A RS 21 Oct 2026"`.

### 45. Three-Tier OOS Review, Roles & Signature Hierarchy Protocol (OOS 报告三级签名流转法定层级与角色绝对清晰铁律 - 🚨 终极锁定)
**CRITICAL**: Strictly distinguish between the direct Microbiology Supervisor and the independent Quality Assurance (QA) Department:
1. **Tier 1: Prepared by (Initiating Analyst / Author)**:
   - **Person**: **Qiyue Chen (QYC)** (Microbiologist).
   - **Action**: Gathers data, drafts narrative, completes form, signs Prepared by: (Signature_0).
2. **Tier 2: Reviewed by (Microbiology Lab Supervisor)**:
   - **Person**: **Robin Seymour (RS)** (Microbiology Supervisor / User's direct supervisor).
   - **Action**: Reviews technical lab details, edits narrative, disposes Section E (N/A RS [Date]), and signs Reviewed by: (Signature_1) AFTER the author signs.
3. **Tier 3: Approved by (QA / Quality Assurance Department)**:
   - **Department**: **Quality Assurance (QA)** (Independent QA reviewers / Director).
   - **Action**: Final compliance review and disposition after lab internal signatures (Approved by (QA): / Signature_2). **QA has not inspected the document while under Supervisor review!**
4. **Mandatory Routing Sequence**:
   Qiyue Chen (Prepared by) ✍️ -> Robin Seymour (Reviewed by) ✍️ -> QA Department (Approved by) ✍️.


### 46. Mandatory Raw EM Logbook Ground-Truth Audit, Anti-Template Hallucination & Positive Plate Lifecycle Protocol (环境监测原始扫描件地毯式核查、反模板脑补与阳性平板生命周期法定铁律 - 终身基准)

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


## 🚨🚨 RULE 0Z (TOP PRIORITY): Automated Pre-Flight Integrity & Anti-Contradiction Assertion Gate (`assert_no_facility_or_event_contradictions`)

**CRITICAL**: Every OOS report generator and transcription script MUST enforce the automated assertion gate defined in `utils.py`:

1. **Event Count Single-Truth Invariant**:
   - `Text Field49` (Page 3 narrative) event count MUST match Table 1 (`expected_events`) exactly. Zero tolerance for narrative stating 15 while Table 1 states 19.

2. **BSC & Cleanroom Spatial Invariant (Figure 6 Ground Truth)**:
   - `BSC E001938` $	o$ **`Cleanroom 145 (CR145)`** ONLY. Pairing with `Cleanroom 144` is blocked by code.
   - `BSC E001937` $	o$ **`Cleanroom 144 (CR144)`** ONLY. Pairing with `Cleanroom 145` is blocked by code.

3. **Active Air Morphology Defense Invariant**:
   - Prohibits morphology comparison claims when environmental isolates are cocci.
