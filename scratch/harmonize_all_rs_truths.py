# =========================================================================
# 1. Cleanly update .agents/AGENTS.md (RULE 0E, RULE 0K, RULE 0T)
# =========================================================================
with open('.agents/AGENTS.md', 'rb') as f:
    b_ag = f.read()

# Fix 1A: RULE 0E
old_0e = '1. **多套间/多洁净室分行规范**: 在 Cleanroom 认证栏目（如 Page 2 `Text Field32`/`33` 及 Word 表格 `smart_cr_id`），多套间严禁使用 `&` 并列（严禁写成 `E001738 (Suite 116) & E001978 (L-Suite)`）。必须采用标准多行对齐：第一行 `E001738 (CR116)`，第二行 `E001978 (CR144)`，右侧对应校准期 `Dec 2026\\rDec 2026`（字号 4.4pt）。'.encode('utf-8')
new_0e = '1. **多套间/多洁净室分行规范**: 在 Cleanroom 认证栏目（如 Page 2 `Text Field32`/`33` 及 Word 表格 `smart_cr_id`），多套间严禁使用 `&` 并列（严禁写成 `E001738 (Suite 116) & E001978 (L-Suite)`）。必须采用标准多行对齐：例如 L-Suite 双房间填报第一行 `E001978 (CR144)`、第二行 `E001979 (CR145)`，右侧对应校准期 `Nov 2026\\rNov 2026`（字号 5.0pt）；若涉及 CR116 (`E001738`) 则对应 `Dec 2026`。'.encode('utf-8')
assert old_0e in b_ag, "old_0e not found in AGENTS.md"
b_ag = b_ag.replace(old_0e, new_0e)

# Fix 1B: RULE 0K
old_0k = '   - 第一行：`E001738 (CR116)`\r\n\r\n\r\n\r\n   - 第二行：`E001978 (CR144)`'.encode('utf-8')
new_0k = '   - L-Suite 双房间标准（Processing in CR145 + Changeover in CR144）：第一行 `E001978 (CR144)`，第二行 `E001979 (CR145)`，对应校准期均为 `Nov 2026\\rNov 2026`（5.0pt）\r\n\r\n   - 跨套间标准（如 CR116 + CR144）：`E001738 (CR116)`（校准期 `Dec 2026`）与 `E001978 (CR144)`（校准期 `Nov 2026`）'.encode('utf-8')
assert old_0k in b_ag, "old_0k not found in AGENTS.md"
b_ag = b_ag.replace(old_0k, new_0k)

# Fix 1C: RULE 0T
old_0t = '3. **L-Suite (Suite 144 / 145) Operational Invariant**:\r\n   - **Cleanroom 144 (ISO 7 核心室 / Scan RDI Processing BSC E001938 & Changeover BSC E001937)**:\r\n     - Regular Operating Analysts: **ELB (Ellen Bennett)**, **KSM (Karla Silva)**, **VV (Varsha Subramanian)**, **EN (Elysse Nioupin)**.'.encode('utf-8')
new_0t = '3. **L-Suite (Suite 145 / 144) Operational Invariant (RS Verified Ground Truth)**:\r\n   - **Cleanroom 145 (`CR145 / E001979`, Cert Due `Nov 2026` — ISO 7 核心室 / Scan RDI Processing BSC `E001938`, Cal Due `Jul 2027`)**:\r\n     - Primary Sample Introduction, Pre-labeling & Filtration Processing.\r\n   - **Cleanroom 144 (`CR144 / E001978`, Cert Due `Nov 2026` — ISO 7 缓冲室 / Scan RDI Changeover BSC `E001937`, Cal Due `Jul 2027`)**:\r\n     - Changeover & Sanitization Procedures.\r\n   - Regular Operating Analysts: **ELB (Ellen Bennett)**, **KSM (Karla Silva)**, **VV (Varsha Subramanian)**, **EN (Elysse Nioupin)**.'.encode('utf-8')
assert old_0t in b_ag, "old_0t not found in AGENTS.md"
b_ag = b_ag.replace(old_0t, new_0t)

with open('.agents/AGENTS.md', 'wb') as f:
    f.write(b_ag)
print('Successfully harmonized RULE 0E, 0K, 0T in .agents/AGENTS.md')

# =========================================================================
# 2. Cleanly update PRD.md (Mandate 5, Mandate 11, Rule 59, Rule 62)
# =========================================================================
with open('PRD.md', 'rb') as f:
    b_prd = f.read()

# Fix 2A: Mandate 5 (RULE 0E)
old_prd_0e = '   - **多套间/多洁净室分行规范**: 在 Cleanroom 认证栏目（如 Page 2 `Text Field32`/`33` 及 Word 表格 `smart_cr_id`），多套间严禁使用 `&` 并列（严禁写成 `E001738 (Suite 116) & E001978 (L-Suite)`）。必须采用标准多行对齐：第一行 `E001738 (CR116)`，第二行 `E001978 (CR144)`，右侧对应校准期 `Dec 2026\\rDec 2026`（字号 4.4pt）。'.encode('utf-8')
new_prd_0e = '   - **多套间/多洁净室分行规范**: 在 Cleanroom 认证栏目（如 Page 2 `Text Field32`/`33` 及 Word 表格 `smart_cr_id`），多套间严禁使用 `&` 并列（严禁写成 `E001738 (Suite 116) & E001978 (L-Suite)`）。必须采用标准多行对齐：例如 L-Suite 双房间填报第一行 `E001978 (CR144)`、第二行 `E001979 (CR145)`，右侧对应校准期 `Nov 2026\\rNov 2026`（字号 5.0pt）；若涉及 CR116 (`E001738`) 则对应 `Dec 2026`。'.encode('utf-8')
assert old_prd_0e in b_prd, "old_prd_0e not found in PRD.md"
b_prd = b_prd.replace(old_prd_0e, new_prd_0e)

# Fix 2B: Mandate 11 (RULE 0K)
old_prd_0k = '      - 第一行：`E001738 (CR116)`\r\n\r\n\r\n\r\n      - 第二行：`E001978 (CR144)`'.encode('utf-8')
new_prd_0k = '      - L-Suite 双房间标准（Processing in CR145 + Changeover in CR144）：第一行 `E001978 (CR144)`，第二行 `E001979 (CR145)`，对应校准期均为 `Nov 2026\\rNov 2026`（5.0pt）\r\n\r\n      - 跨套间标准（如 CR116 + CR144）：`E001738 (CR116)`（校准期 `Dec 2026`）与 `E001978 (CR144)`（校准期 `Nov 2026`）'.encode('utf-8')
assert old_prd_0k in b_prd, "old_prd_0k not found in PRD.md"
b_prd = b_prd.replace(old_prd_0k, new_prd_0k)

# Fix 2C: Rule 59 (OOS-262289)
old_prd_59 = '''        - 测试方法与设备: Scan RDI Sterility Test (`MICRO-SOP-12`, `ENG-SOP-4`), Instrument `E001040`, Cleanroom 144 (L-Suite), Processing BSC `E001938`, Changeover BSC `E001937`.\r\n        - 结果: 15 total events, 4 confirmed microbial events exhibiting `Cocci shaped morphology`.'''.encode('utf-8')
new_prd_59 = '''        - 测试方法与设备: Scan RDI Sterility Test (`MICRO-SOP-12`, `ENG-SOP-4`), Instrument `E001040` (Cal Due `Oct 2026`), L-Suite Cleanroom 145 (`E001979`, Cert Due `Nov 2026`, Processing BSC `E001938` Cal Due `Jul 2027`) & Cleanroom 144 (`E001978`, Cert Due `Nov 2026`, Changeover BSC `E001937` Cal Due `Jul 2027`).\r\n        - 结果 (RS Reviewed): **19 total events, 4 confirmed microbial events** exhibiting `Cocci shaped morphology` (verified per Scan Data packet).'''.encode('utf-8')
assert old_prd_59 in b_prd, "old_prd_59 not found in PRD.md"
b_prd = b_prd.replace(old_prd_59, new_prd_59)

old_prd_59_t1 = '| Ellen Bennett | Varsha Subramanian | [ETX-260929-0931](https://etrax.eagleanalytical.com/Submission/Details/sM81r08$r4X9k5LrerlJ7g__) | 15 | 4 | Cocci shaped morphology |'.encode('utf-8')
new_prd_59_t1 = '| Ellen Bennett | Varsha Subramanian | [ETX-260929-0931](https://etrax.eagleanalytical.com/Submission/Details/sM81r08$r4X9k5LrerlJ7g__) | 19 | 4 | Cocci shaped morphology |'.encode('utf-8')
assert old_prd_59_t1 in b_prd, "old_prd_59_t1 not found in PRD.md"
b_prd = b_prd.replace(old_prd_59_t1, new_prd_59_t1)

# Fix 2D: Rule 62 (L-Suite Operational Invariant)
old_prd_62 = '        3. **L-Suite (Suite 144 / 145) 常态化排班与工位**:\r\n           - **Cleanroom 144 (ISO 7 核心室 / Scan RDI 处理 BSC E001938 & 换台 BSC E001937)**:\r\n             - 常态化操作分析员：**ELB (Ellen Bennett)**、**KSM (Karla Silva)**、**VV (Varsha Subramanian)**、**EN (Elysse Nioupin)**。'.encode('utf-8')
new_prd_62 = '        3. **L-Suite (Suite 145 / 144) 常态化排班与工位 (RS Verified Ground Truth)**:\r\n           - **Cleanroom 145 (`CR145 / E001979`, Cert Due `Nov 2026` — ISO 7 核心室 / Scan RDI 处理 BSC `E001938`, Cal Due `Jul 2027`)**:\r\n             - 专职样品引入、预贴签与过滤处理（Processing Analyst）。\r\n           - **Cleanroom 144 (`CR144 / E001978`, Cert Due `Nov 2026` — ISO 7 缓冲室 / Scan RDI 换台清场 BSC `E001937`, Cal Due `Jul 2027`)**:\r\n             - 专职换批消毒清场（Changeover Analyst）。\r\n           - 常态化操作分析员：**ELB (Ellen Bennett)**、**KSM (Karla Silva)**、**VV (Varsha Subramanian)**、**EN (Elysse Nioupin)**。'.encode('utf-8')
assert old_prd_62 in b_prd, "old_prd_62 not found in PRD.md"
b_prd = b_prd.replace(old_prd_62, new_prd_62)

with open('PRD.md', 'wb') as f:
    f.write(b_prd)
print('Successfully harmonized Mandate 5, 11, Rule 59, Rule 62 in PRD.md')
