with open('.agents/AGENTS.md', 'rb') as f:
    b = f.read()

target = b'## SOP: Creating a New OOS Module (Template Handoff Workflow)'
idx = b.find(target)
sep_idx = b.rfind(b'---', 0, idx)

rule_0x_text = '''\r\n\r\n## 🚨🚨 RULE 0X (TOP PRIORITY): Robin Seymour (RS) Quality Review & Defensibility Standards (ScanRDI 案卷与套间布局终极质保规范)\r\n\r\n**CRITICAL**: Across all ScanRDI and cleanroom OOS investigations, all generated forms and narratives MUST strictly incorporate Robin Seymour's verified Quality Standards:\r\n\r\n1. **原始数据包事件数唯一法定原则 (Event Count Source Truth)**:\r\n   - 严禁仅采信分析员手写便签或口头流转数据（Analyst data entries）。**必须严格调取 Scan Data Packet 仪器原始打印报告核验总事件数（Total Events）与确证微生物事件数（Confirmed Microbial Events）**。\r\n   - 例如 OOS-262289 中分析员记录为 15，实际原始数据包显示为 19 total events, 4 confirmed microbial events，以原始数据包 19 为唯一法定依据。Table 1 与正文叙述必须绝对严格一致。\r\n\r\n2. **L-Suite 物理房间真实对应布局 (L-Suite Room Mapping)**:\r\n   - **BSC E001938** 位于 **Cleanroom 145 (CR145)**（用于样品引入、预贴签与过滤处理）。\r\n   - **BSC E001937** 位于 **Cleanroom 144 (CR144)**（用于换批消毒清场 Changeover）。\r\n   - 🚨 **严禁混为一谈**：绝不可将两台 BSC 笼统写在 CR144！必须精准体现各 BSC 所在的真实房间编号。\r\n\r\n3. **双套间认证与校准期精准填报 (Environmental Certification Precision)**:\r\n   - Cleanroom 认证（Page 2 `Text Field32`）：必须同时填报 `E001978 (CR144)\\rE001979 (CR145)`（字号 5.0pt）。\r\n   - 校准有效期（`Text Field33`）：对应 `Nov 2026\\rNov 2026`（字号 5.0pt）。\r\n   - BSC 校准有效期（`Text Field44`）：E001938 与 E001937 均填报 `Jul 2027`。\r\n   - ScanRDI 仪器校准期（`Text Field31`/`35`）：填报 `Oct 2026`。\r\n\r\n4. **微生物形态学对比科学严密性铁律 (Morphology Defensibility Standard)**:\r\n   - 当环境菌（如 *Staphylococcus epidermidis*）同属球菌形态时，**严禁宣称其形态与阳性样本 Cocci 截然不同**！\r\n   - 必须完全删除形态差异对比，将防守论点纯粹建立在“物理套间硬隔离、闭门、正压梯度气流屏障，以及 ISO 5 台面和手套完全无菌（No Growth）”之上。\r\n\r\n5. **月度清洁与交叉污染极简精炼句式 (Streamlined Quality Phrasing)**:\r\n   - 月度清洁：`"It was documented that all chemical indicators passed, which confirmed the efficient monthly cleaning of all cleanrooms in L-Suite."`\r\n   - 同日阳性：`"One other sample (ETX-XXXXXX-XXXX, OOS-XXXXXX) yielded a positive result on [Date]. However, that sample was processed by a different processing analyst during a different shift. The two testing events were segregated by shift, processing analyst and primary containment equipment."`\r\n\r\n'''.encode('utf-8')

b_new = b[:sep_idx] + rule_0x_text + b[sep_idx:]
with open('.agents/AGENTS.md', 'wb') as f:
    f.write(b_new)
print('AGENTS.md updated with RULE 0X cleanly.')
