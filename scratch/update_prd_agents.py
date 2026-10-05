import os

rule_text_prd = '''    23. *Celsis Narrative Precision & 6-Month Analyte History Standard (Celsis 叙述精简与特定分析物历史规范):*
        *   **Omission of Negative TSB Numerical Cutoff Sentence (阴性管数值细节精简):**
            - In Paragraph 8 (Celsis RLU reading summary), do NOT include detailed negative RLU and cutoff numbers for the negative TSB bottle (e.g. omit: "The corresponding TSB sample container ETX-XXXXXX-XXXX tested negative with X RLU, below the TSB cutoff of X RLU and negative control of X RLU.").
            - Simply state: "All other FTM and TSB sample bottles in the same batch tested negative."
        *   **Omission of Lot Retesting Submission Sentence (批号复检历史段落剔除):**
            - Do NOT include generic statements regarding whether the lot was submitted for retesting (e.g. omit: "A review of the lot history shows that there have been no additional submissions of sample lot X for retesting..."). Omit this paragraph entirely.
        *   **Locked Standard Phrasing for 6-Month Analyte History (6个月特定分析物历史标准句式):**
            - Standard template syntax:
              `Analyzing a 6-month sample history for [Client Name] indicates, this specific analyte "[Analyte / Sample Name]" has had no prior failures using the Celsis Sterility testing during this period.`
            - Concrete example:
              `Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "MOTs-C 10 MG/ML (5 ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.`
'''

with open('PRD.md', 'r', encoding='utf-8') as f:
    prd_lines = f.readlines()

for i, l in enumerate(prd_lines):
    if '## Pending/Future Work' in l:
        prd_lines.insert(i, rule_text_prd)
        break

with open('PRD.md', 'w', encoding='utf-8') as f:
    f.writelines(prd_lines)
print('Updated PRD.md')

rule_text_agents = '''
### 18. Celsis Narrative Precision & 6-Month Analyte History Standard (Celsis 叙述精简与特定分析物历史规范)
**CRITICAL**: Across all Celsis Sterility OOS investigation reports:
1. **Omission of Negative TSB Numerical Cutoff Sentence (阴性管数值细节精简)**:
   - In Paragraph 8 (Celsis RLU reading summary), do NOT report the granular negative RLU and cutoff numbers for the negative TSB bottle (e.g. omit: `"The corresponding TSB sample container ETX-XXXXXX-XXXX tested negative with X RLU, below the TSB cutoff of X RLU and negative control of X RLU."`).
   - Simply state: `"All other FTM and TSB sample bottles in the same batch tested negative."`
2. **Omission of Lot Retesting Submission Sentence (批号复检历史段落剔除)**:
   - Do NOT include generic statements regarding whether the lot was submitted for retesting (e.g. omit: `"A review of the lot history shows that there have been no additional submissions of sample lot X for retesting..."`). Omit this paragraph entirely.
3. **Locked Standard Phrasing for 6-Month Analyte History (6个月特定分析物历史标准句式)**:
   - Standard template syntax:
     `Analyzing a 6-month sample history for [Client Name] indicates, this specific analyte "[Analyte / Sample Name]" has had no prior failures using the Celsis Sterility testing during this period.`
   - Concrete example:
     `Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "MOTs-C 10 MG/ML (5 ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.`
'''

with open('.agents/AGENTS.md', 'r', encoding='utf-8') as f:
    agents_text = f.read()

agents_text = agents_text.rstrip() + '\n' + rule_text_agents

with open('.agents/AGENTS.md', 'w', encoding='utf-8') as f:
    f.write(agents_text)
print('Updated AGENTS.md')
