import os, sys

def update_prd():
    with open('PRD.md', 'r', encoding='utf-8') as f:
        content = f.read()

    target = '### Rule 86 / RULE 0AD: Day 18+ Subculture Environmental Monitoring (Table 3) & Complete Traceability Standard'
    idx = content.find(target)
    assert idx != -1, 'Target not found in PRD.md'

    new_rule_86 = '''### Rule 86 / RULE 0AD: Day 18+ Subculture Environmental Monitoring (Table 3), GDP Sign-off & Complete Traceability Golden Standards (次培养 Day 18+ 专属环境监测附表、GDP 签署与全流程防守十大铁律 - 🚨 主管 Robin Seymour 权威审阅终极版)

**CRITICAL**: Based on Microbiology Lab Supervisor Robin Seymour (RS) formal review of `OOS-262329` (GoGoMeds Select, Sermorelin 0.9mg/ml, Lot 17009) on 05-Nov-2026 across all sterility investigators (@Dhvanir Kansara, @Olugbenga Ajayi, @Qiyue Chen):
Whenever a sterility sample involves Day 14 formulation turbidity, inconclusive reading, and subsequent aseptic subculture into fresh media (Day 18+ subculture), all future OOS reports MUST strictly adhere to the following 10 golden rules to achieve zero supervisory revisions:

1. **次培养必须附带 Table 3 专属环境监测表 (Mandatory Table 3 for Subculture Event)**:
   - 凡涉及 Day 14 混浊、初读 Inconclusive 并进行次培养（Day 18+ subculture）的样品，次培养由次培养分析员（如 Ayomide Odugbesi / AOD）在另一天（如 07Oct26）开展独立无菌操作。
   - 必须在报告中新增 **Table 3**（`Table 3: Environmental Monitoring from Day 18+ Subculture Performed on [DDMMMYY]`），涵盖：
     - **Personnel EM**: 次培养分析员当天的左右手手指触碟（Date of Testing，No Growth）。
     - **BSC EM**: 次培养所用 BSC（如 ISO 5 BSC E001314）当日表面与沉降碟（Date of Testing，No Growth）。
     - **Weekly Active Air & Surface Sampling**: 次培养当周该洁净套间（CR115）的浮游菌与表面监测。
   - 无论是正式调查表 PDF 还是独立 Word 表格文件（`Tables OOS-...docx`），都必须完整收录 Table 1、Table 2、Table 3！

2. **洁净室浮游菌检出透明披露与缓冲间/正压物理隔离技术防守 (Subculture EM Recovery & Physical Segregation Rationale)**:
   - 若次培养当周洁净室浮游菌有检出（例如：06Oct26 Buffer Room 115A 检出 1 CFU *Corynebacterium segmentosum* under ETX-261013-0297）：
     - 必须在 Table 3 中如实透明列出并附带 EagleTrax 超链接。
     - 正文中必须严密防守：明确说明该分离菌（*Corynebacterium segmentosum*）与产品污染菌（*Bacillus megaterium*）属于完全不同的微生物（形态与种属均不符）。
     - 明确指出 Room 115 为 ISO 8 前室，Room 115A 为 ISO 7 缓冲室，两者与最内侧的 ISO 7 核心操作室 115B（内含 ISO 5 BSC E001314）之间均有气闸门与阶梯正压保护。
     - 核心区 115B、两位分析员手套及 BSC 表面均零检出（No Growth），客观证明外围杂菌未穿透至无菌操作环境。

3. **次培养天数严格数学相减法则 (Subculture Day Count Mathematical Consistency Rule)**:
   - 次培养分装管培养天数计算公式为：**原始瓶总培养天数 - 次培养执行/初读天数**。
   - 示例：原始瓶第 18 天（Day 18），第 14 天（Day 14）执行次培养，则次培养管的培养天数为 18 - 14 = 4（**Day 4 of the subculture aliquots**）！
   - 🚨 **严禁写成 Day 5**！所有叙述中必须严格核算天数，保持数学绝对严谨。

4. **Section E 规范化 GDP 签署铁律 (`N/A QYC [DDMMMYY]` Sign-Off Rule)**:
   - 在官方调查表（CORP-FORM-21 / CORP-FORM-22）中，所有不适用的可填文本框（Fillable Text Fields）中**严禁仅填写裸露的 `N/A`**！
   - 必须严格遵循良好文件规范（GDP），签署：**`N/A [调查员缩写] [DDMMMYY]`**（例如：`N/A QYC 04Nov26`）！
   - 在 Phase II Section E（纠正与预防措施）中必须全量覆盖：
     - `Corrective Action(s)`: `N/A QYC 04Nov26`
     - `Preventive Action(s)`: `N/A QYC 04Nov26`
     - `Other: ____________`: `N/A QYC 04Nov26`
     - `CAPA required (Yes, CAPA #)`: `N/A QYC 04Nov26`
     - 对应勾选框：`External Phenomena` = Yes，`CAPA required -> No` = Yes，`FOLLOW-UP REQUIRED -> No` = Yes。

5. **初检与复检阶段 Table 1 样品范围严格区隔 (Phase I vs Phase II Table 1 Scope Separation)**:
   - **Phase I Table 1**: 仅收录初检受调查样品（`ETX-260624-0251 (BUD-T90)`），**严禁在 Phase I 提前混入 Phase II 复检样品**！
   - **Phase II Table 1**: 完整统筹呈现初检与复检双样品（`ETX-260624-0251` Fail vs `ETX-261021-0445` Pass）。

6. **套间主备设备与安全柜全量覆盖溯源 (Comprehensive Primary & Secondary Equipment Traceability)**:
   - Page 1 / Page 2 设备栏目不仅要列出初检培养箱（E001356, E001357 与传感器 E001450, E001449），还必须完整列出备份与关联培养箱（`Incubator E001034 (Monitored by Sensor E001501)` 与 `Incubator E001031 (Monitored by Sensor E001505)`）。
   - 必须显式在设备栏中注明初级工程控制设备 `ISO 5 BSC E001314`（Cal Due Jun 2027）。
   - 在 Phase II 表单 Page 1 的次级设备栏中，严禁填 `Not Applicable`，必须完整填入 `Incubator E001034, Sensor E001501\\rIncubator E001031, Sensor E001505` 与校准期 `Aug 2027, Feb 2027\\rAug 2027, Feb 2027`。

7. **Day 14 初读日期严密对齐 (Timeline Synchronization)**:
   - 混浊初读与次培养执行必须严格记录为同一日（`On 07Oct26 (Day 14 read)...`），杜绝 08Oct26 与 07Oct26 的前后矛盾。

8. **微生物无菌专有表单勾选绝对规范 (Microbiology Checkbox Invariants)**:
   - Phase I Page 3 `Check Box79`（色谱进样/复溶相关）：必须勾选 `Yes`（`N/A`），因为该项专指理化/HPLC 色谱分析，不适用于无菌直接接种。
   - Phase II Page 2 `Check Box42`（配方工单 Formula Worksheet）：必须勾选 `Yes`（对应 `No`），因为无菌检验属于实验室检验规程，不提供制剂配方工单。

9. **环境监测叙述大写分节规范 (Narrative Compartmentalization Standard)**:
   - Section D 环境监测叙述中，必须使用清晰的标准全大写标题分节：
     - `ENVIRONMENTAL MONITORING (EM) RESULTS:`
     - `ISO 5 BSC and Personnel EM results for Initial Processing - [DDMMMYY]:`
     - `ISO 5 BSC and Personnel EM results for Subculture - [DDMMMYY]:`
     - `Cleanroom EM Results for Initial Processing - [DDMMMYY]:`
     - `Cleanroom EM Results for Subculture- [DDMMMYY]:`

10. **调查结论强化与 BUD 稳定性研究结案承接句式 (Expanded Investigation Conclusion & BUD Transition)**:
    - Phase I 结尾必须强化调查结论，明确排除实验室人为失误，并将 Phase II 复检紧密锚定在客户针对 BUD-T90 稳定性研究的专项申请上：
      `"Consequently, the Phase I laboratory investigation did not identify any assignable laboratory cause contributing to the initial failure, and the original result is considered valid. At the request of the client, additional vials from the same lot were submitted for USP <71> sterility retesting (Phase II) to evaluate the sterility of the lot for the BUD-T90 stability time point."`

11. **主管审核闭环流转流程 (Supervisor Review & Signature Route)**:
    - 主管 Robin Seymour 审阅并在原件上批注修改后，文件命名为 `OOS-... - Phase I - Reviewed - QYC.pdf`。
    - 调查员（Qiyue Chen）复核无误后，只需在 `Signature_0` (Prepared by) 处完成证书数字签名，即可直接回复邮件流转给 Robin 与 QA 签署终审！
'''

    updated_content = content[:idx] + new_rule_86
    with open('PRD.md', 'w', encoding='utf-8') as f:
        f.write(updated_content)
    print('Updated PRD.md successfully.')

    # Also update AGENTS.md
    with open('.agents/AGENTS.md', 'r', encoding='utf-8') as f:
        agents_content = f.read()
    
    idx_agents = agents_content.find('### Rule 86')
    if idx_agents != -1:
        updated_agents = agents_content[:idx_agents] + new_rule_86
        with open('.agents/AGENTS.md', 'w', encoding='utf-8') as f:
            f.write(updated_agents)
        print('Updated .agents/AGENTS.md successfully.')

if __name__ == '__main__':
    update_prd()
