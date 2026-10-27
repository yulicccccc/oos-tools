import os

# 1. Update .agents/AGENTS.md
with open('.agents/AGENTS.md', 'rb') as f:
    agents_bytes = f.read()

rule_0y_text = """

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
"""

target = b'## SOP: Creating a New OOS Module (Template Handoff Workflow)'
if target in agents_bytes:
    new_bytes = agents_bytes.replace(target, rule_0y_text.encode('utf-8') + b'\n\n---\n\n' + target)
    with open('.agents/AGENTS.md', 'wb') as f:
        f.write(new_bytes)
    print('Successfully inserted RULE 0Y into .agents/AGENTS.md')
else:
    print('Target not found in .agents/AGENTS.md')

# 2. Update PRD.md
with open('PRD.md', 'rb') as f:
    prd_bytes = f.read()

rule_71_text = """

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
"""

with open('PRD.md', 'ab') as f:
    f.write(rule_71_text.encode('utf-8'))
print('Successfully appended Rule 71 to PRD.md')
