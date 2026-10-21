import os

# =========================================================================
# 1. Update PRD.md
# =========================================================================
with open('PRD.md', 'rb') as f:
    prd_data = f.read()

rule_62_prd = '''62. **Good Documentation Practices (GDP) N/A Blank Field Sign-Off Standard (N/A + 分析员官方缩写 + 2位年份日期 DDMMMYY 规范):**
    *   **适用范围**: 全局适用于所有 OOS 调查模块（ScanRDI, Celsis, USP <71>, Environmental Monitoring）及官方表单（Phase I CORP-FORM-21, Phase II CORP-FORM-22 等）。凡任一文本框、栏目、表格单元格或溢出续页（如 Page 3 Section C "Other: ________" Text Field48、Page 5 "Phase I Summary (continued):" Text Field51 等）无需填写、不适用或划消签注时，必须严格执行该规范。
    *   **法定格式**: `N/A <Analyst Initials> <DDMMMYY>`（例如：`N/A QYC 21Oct26`）。
    *   **2 位年份绝对铁律（Strict 2-Digit Year Mandate）**: 日期部分必须且只能采用 2 位年份（如 `21Oct26`、`01Oct26`、`15Sep26`），**🚨 绝对严禁使用 4 位年份（严禁写成 `21Oct2026`）**！
    *   **分析员官方花名册对应（Analyst Initials Compliance）**: 缩写必须严格对应官方权威花名册（如 Qiyue Chen 必须为 `QYC`，严禁写成 `QC` 或 `KC`，遵守 RULE 0A）。
    *   **空格与标点铁律**: 各字段间以单个空格分隔，日期中不得包含斜杠或连字符，句尾严禁加句号。


'''.encode('utf-8')

target_prd = b'## Pending/Future Work'
if b'Good Documentation Practices (GDP) N/A Blank Field Sign-Off Standard' not in prd_data:
    idx = prd_data.find(target_prd)
    if idx != -1:
        new_prd_data = prd_data[:idx] + rule_62_prd + prd_data[idx:]
        with open('PRD.md', 'wb') as f:
            f.write(new_prd_data)
        print('PRD.md successfully updated with Rule 62.')
    else:
        print('Target not found in PRD.md!')
else:
    print('Rule 62 already in PRD.md.')

# =========================================================================
# 2. Update .agents/AGENTS.md
# =========================================================================
with open('.agents/AGENTS.md', 'rb') as f:
    agents_data = f.read()

rule_0q_agents = '''

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
'''.encode('utf-8')

target_agents = b'## SOP: Creating a New OOS Module (Template Handoff Workflow)'
if b'RULE 0Q (TOP PRIORITY)' not in agents_data:
    idx = agents_data.find(target_agents)
    if idx != -1:
        sep_idx = agents_data.rfind(b'---', 0, idx)
        if sep_idx != -1:
            new_agents_data = agents_data[:sep_idx] + rule_0q_agents + b'\n\n' + agents_data[sep_idx:]
            with open('.agents/AGENTS.md', 'wb') as f:
                f.write(new_agents_data)
            print('.agents/AGENTS.md successfully updated with RULE 0Q.')
        else:
            print('Separator not found before target in AGENTS.md!')
    else:
        print('Target not found in AGENTS.md!')
else:
    print('RULE 0Q already in AGENTS.md.')

# =========================================================================
# 3. Update PROJECT_STATE.md
# =========================================================================
with open('PROJECT_STATE.md', 'rb') as f:
    state_data = f.read()

state_addition = b'- **Good Documentation Practices (GDP) N/A Sign-Off Standard (RULE 0Q & PRD Rule 62):** Codified canonical GDP N/A sign-off format across all OOS modules and forms: `N/A <Analyst Initials> <DDMMMYY>` with strict 2-digit year (e.g., `N/A QYC 21Oct26`, strictly prohibiting 4-digit years like `21Oct2026`). Corrected Page 3 `Text Field48` and Page 5 `Text Field51` in OOS-262289 master packet.\n'

if b'Good Documentation Practices (GDP) N/A Sign-Off Standard' not in state_data:
    target_state = b'### What Was Completed\n'
    idx = state_data.find(target_state)
    if idx != -1:
        new_state_data = state_data[:idx + len(target_state)] + state_addition + state_data[idx + len(target_state):]
        with open('PROJECT_STATE.md', 'wb') as f:
            f.write(new_state_data)
        print('PROJECT_STATE.md successfully updated.')
    else:
        target_state_crlf = b'### What Was Completed\r\n'
        idx = state_data.find(target_state_crlf)
        if idx != -1:
            new_state_data = state_data[:idx + len(target_state_crlf)] + state_addition + state_data[idx + len(target_state_crlf):]
            with open('PROJECT_STATE.md', 'wb') as f:
                f.write(new_state_data)
            print('PROJECT_STATE.md successfully updated (CRLF).')
else:
    print('State already recorded in PROJECT_STATE.md.')
