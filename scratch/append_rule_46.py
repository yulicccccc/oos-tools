import os

rule_text = """

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
"""

with open('.agents/AGENTS.md', 'a', encoding='utf-8') as f:
    f.write(rule_text)
print('Appended Rule 46 to AGENTS.md')

with open('PRD.md', 'r', encoding='utf-8') as f:
    prd_content = f.read()

insert_idx = prd_content.find('## Pending/Future Work')
if insert_idx != -1:
    prd_updated = prd_content[:insert_idx] + '\n*   **Mandatory Raw EM Logbook Ground-Truth Audit & Positive Plate Lifecycle Protocol (Rule 46 / Rule 65):**' + rule_text.replace('### 46.', '') + '\n\n' + prd_content[insert_idx:]
    with open('PRD.md', 'w', encoding='utf-8') as f:
        f.write(prd_updated)
    print('Inserted Rule 46 into PRD.md')
else:
    with open('PRD.md', 'a', encoding='utf-8') as f:
        f.write(rule_text)
    print('Appended Rule 46 to PRD.md')
