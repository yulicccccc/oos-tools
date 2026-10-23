# scratch/insert_rule0v.py
import os

agents_path = r".agents\AGENTS.md"
prd_path = r"PRD.md"

rule_agents = """## 🚨🚨 RULE 0V (TOP PRIORITY): Weekly Environmental Monitoring (EM) "Week of Testing" Calendar Alignment Protocol (周度环境监测 "Week of Testing" 自然周严格对齐铁律)

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

"""

with open(agents_path, 'r', encoding='utf-8') as f:
    agents_content = f.read()

target = "## SOP: Creating a New OOS Module (Template Handoff Workflow)"
if target in agents_content:
    agents_content = agents_content.replace(target, rule_agents + target)
    with open(agents_path, 'w', encoding='utf-8') as f:
        f.write(agents_content)
    print("Successfully added RULE 0V to .agents/AGENTS.md")
else:
    print("Target not found in AGENTS.md")

rule_prd = """67. **Weekly Environmental Monitoring (EM) "Week of Testing" Calendar Alignment Protocol (周度环境监测 "Week of Testing" 自然周严格对齐铁律 - 🚨 终极锁定):**
    *   **核心法则**: 在所有无菌检验调查（Celsis、ScanRDI、USP <71>）附表 Table 2 & Table 3 中：
        - `Week of Testing`（检验当周）必须严格与测试日期所在的同一个 Monday-to-Sunday 自然周对齐！
        - **严禁整周错位**：若检验在周一进行，当周的周检取当周中（如周四）采样的记录，严禁将上一周（T-1 周）的周检记录错填为检验当周。
        - **官方先例基准**:
          - `OOS-262017`: 24Aug26 (Mon) Processing ➡️ Table 2 为 25Aug26 (Tue)；31Aug26 (Mon) Aliquoting ➡️ Table 3 为 04Sep26 (Fri)。
          - `OOS-262333`: 05Oct26 (Mon) Processing ➡️ Table 2 为 08Oct26 (Thu)；12Oct26 (Mon) Aliquoting ➡️ Table 3 为 15Oct26 (Thu)。
"""

with open(prd_path, 'r', encoding='utf-8') as f:
    prd_content = f.read()

if "67. **Weekly Environmental Monitoring" not in prd_content:
    with open(prd_path, 'a', encoding='utf-8') as f:
        f.write("\n\n" + rule_prd)
    print("Successfully added Section 67 to PRD.md")
else:
    print("Section 67 already in PRD.md")
