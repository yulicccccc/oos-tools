import os

script_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = os.path.abspath(os.path.join(script_dir, ".."))

agents_file = os.path.join(root_dir, ".agents", "AGENTS.md")
prd_file = os.path.join(root_dir, "PRD.md")
state_file = os.path.join(root_dir, "PROJECT_STATE.md")

new_rules_agents = """

## 🚨 Laboratory Physical Topology, EM Dual-Track Tracing & cGMP Investigation Golden Rules

### 1. Physical Spatial Topology Chain (物理空间拓扑链与房间推导法则)
- **Do not blindly accept scan documents without spatial validation.** Every BSC is physically situated in a specific cleanroom, which belongs to a specific cleanroom suite:
  $$\\text{BSC 1316} \\longrightarrow \\text{Cleanroom 114B (ISO 7 核心室)} \\longrightarrow \\text{Cleanroom 114A (ISO 7 缓冲室)} \\longrightarrow \\text{Anteroom 114 (ISO 8 走廊)} \\longrightarrow \\text{Suite 114}$$
  $$\\text{BSC 1798} \\longrightarrow \\text{Cleanroom 114A (ISO 7 缓冲室)} \\longrightarrow \\text{Anteroom 114 (ISO 8 走廊)} \\longrightarrow \\text{Suite 114}$$
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
"""

# 1. Update AGENTS.md
with open(agents_file, "r", encoding="utf-8") as f:
    agents_content = f.read()

if "Laboratory Physical Topology" not in agents_content:
    with open(agents_file, "w", encoding="utf-8") as f:
        f.write(agents_content.rstrip() + "\n" + new_rules_agents)
    print("Updated .agents/AGENTS.md")
else:
    print(".agents/AGENTS.md already has section")

# 2. Update PRD.md
prd_addition = """    *   **2026-09-25 Update (OOS-262080 Full Celsis Production Package & Ground Truth Audit):** Finalized the production pipeline for Celsis OOS-262080 (ETX-260828-0527, Optimal Balance Pharmacy). Enforced 5 cGMP investigation rules: (1) Spatial Topology verification (BSC 1316 in 114B / Suite 114 strictly requires Suite 114 weekly EM; rejecting Suite 115 trap); (2) Dual-Track EM Tracing (Personnel EM follows CCD personal shifts on 03Sep/08Sep/15Sep, Settling & Surface follow BSC 1798 hood history on 07Sep ALA/08Sep CCD/09Sep ALA); (3) Raw Bench Records Precedence (Cuong Du / CCD designated as Aliquoting Analyst based on physical logbook signatures); (4) Dual-Week ISO 8 Anteroom 1 CFU defense integration (04Sep26 ISS under ETX-260914-0487 and 10Sep26 SMO under ETX-260921-0520 defended via physical segregation and 0 CFU critical zone recovery); (5) Post-processing defensive cell overwrite for hardcoded template text. All 5 deliverables generated and verified on Desktop."""

with open(prd_file, "r", encoding="utf-8") as f:
    prd_content = f.read()

if "2026-09-25 Update" not in prd_content:
    target_prd = "producing a unified 7-page official OOS PDF report."
    if target_prd in prd_content:
        prd_content = prd_content.replace(target_prd, target_prd + "\n" + prd_addition)
        with open(prd_file, "w", encoding="utf-8") as f:
            f.write(prd_content)
        print("Updated PRD.md")

# 3. Update PROJECT_STATE.md
state_addition = """- Successfully finalized Celsis OOS-262080 (ETX-260828-0527, Optimal Balance Pharmacy) complete package, resolving all 4 hidden traps in scan files (Suite 115 trap vs Suite 114 ground truth, ISO 8 114 weekly air 1 CFU recovery under ETX-260914-0487 on 04Sep26 by ISS, Aliquoting Analyst attribution to Cuong Du / CCD, and complete BSC 1316/1798 surface logs). Produced and synced 5 clean deliverables: Main DOCX, Standalone Tables DOCX, Main PDF, Standalone Tables PDF, and Complete Combined PDF package.
- Documented and locked the 5 core cGMP investigation rules into `.agents/AGENTS.md` and `PRD.md` (Physical Spatial Topology Chain, EM Dual-Track Tracing, Raw Bench Records Precedence, Zero Hallucination & Sub-threshold Recovery Defense, and Defensive Post-Processing)."""

with open(state_file, "r", encoding="utf-8") as f:
    state_content = f.read()

if "OOS-262080" not in state_content:
    target_state = "Appended the standalone Table PDF as Page 7 to create the unified 7-page official OOS investigation PDF report."
    if target_state in state_content:
        state_content = state_content.replace(target_state, target_state + "\n" + state_addition)
        with open(state_file, "w", encoding="utf-8") as f:
            f.write(state_content)
        print("Updated PROJECT_STATE.md")

print("All documentation updates completed!")
