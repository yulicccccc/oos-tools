# Pin a compact summary of RULE 0P, 0Q, 0X, 0Y (and PRD Rules 61, 62, 70, 71)
# at the very top of .agents/AGENTS.md and PRD.md so they are NEVER truncated
# by the system prompt byte limit or 800-line read_file limit.

top_banner_agents = """# Project-Scoped Rules for Eagle Analytical OOS Tools

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
>    - **Page 2 Cleanroom Fields (`Text Field32`/`33`)**: When testing spans CR145 & CR144, list BOTH: `E001978 (CR144)\\rE001979 (CR145)` (`5.0pt`) and `Nov 2026\\rNov 2026` (`5.0pt`).
> 4. **Active Air Morphology Comparison Ban (`RULE 0X.4`)**: **NEVER** claim environmental isolates (e.g., *Staphylococcus epidermidis*) have "morphology distinctly different from cocci" (since *Staphylococcus* IS cocci-shaped!). Rely strictly on physical room/door segregation, positive pressure cascade, and zero ISO 5/glove growth.
> 5. **RS Streamlined Phrasing (`RULE 0X.5`)**:
>    - Monthly Cleaning: `"It was documented that all chemical indicators passed, which confirmed the efficient monthly cleaning of all cleanrooms in L-suite."`
>    - Concurrent Positive: `"One other sample (ETX-XXXXXX-XXXX, OOS-XXXXXX) yielded a positive result on [DDMMMYY]. However, that sample was processed by a different processing analyst during a different shift. The two testing events were segregated by shift, processing analyst and primary containment equipment."` (Do NOT include third-party client/product/analyst names; always ask the user if there were any other positives on that day!).
> 6. **GDP 2-Digit Year N/A Sign-Off (`RULE 0Q`)**: Always use 2-digit year `DDMMMYY` (`N/A QYC 21Oct26`), NEVER 4-digit year (`21Oct2026`).
> 7. **Master Table Golden Standards (`RULE 0P`)**: 4 canonical banners without dates, `and` instead of `&`, `No Growth` / `Not Applicable` / `None`, omit Gram staining when Microbial ID exists, empty line between multiple species, uniform `7.0pt` font across all cells & hyperlinks, strict 1-page layout, and 2-file deliverable (`CORP-FORM-21-...pdf` + `Tables OOS-...docx`).
"""

with open('.agents/AGENTS.md', 'rb') as f:
    ab = f.read()

old_header_agents = b'# Project-Scoped Rules for Eagle Analytical OOS Tools'
assert old_header_agents in ab, "Header not found in .agents/AGENTS.md"
ab_new = ab.replace(old_header_agents, top_banner_agents.encode('utf-8'), 1)
with open('.agents/AGENTS.md', 'wb') as f:
    f.write(ab_new)
print("Pinned latest updates summary at the very top of .agents/AGENTS.md")

top_banner_prd = """# Product Requirements Document: OOS Tools

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
"""

with open('PRD.md', 'rb') as f:
    pb = f.read()

old_header_prd = b'# Product Requirements Document: OOS Tools'
assert old_header_prd in pb, "Header not found in PRD.md"
pb_new = pb.replace(old_header_prd, top_banner_prd.encode('utf-8'), 1)
with open('PRD.md', 'wb') as f:
    f.write(pb_new)
print("Pinned latest updates summary at the very top of PRD.md")
