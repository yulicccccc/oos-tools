import os

rule = """

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
"""

paths = [
    r"C:\Users\qchen\.gemini\config\AGENTS.md",
    r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\.agents\AGENTS.md"
]

for p in paths:
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            c = f.read()
        if "Side-by-Side Diff Comparison Rule" not in c:
            with open(p, "a", encoding="utf-8") as f:
                f.write(rule)
            print(f"Successfully appended rule to: {p}")
        else:
            print(f"Rule already exists in: {p}")
    else:
        print(f"Path does not exist: {p}")
