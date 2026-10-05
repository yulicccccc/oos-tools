import sys

section19 = """

### 19. Strict 1-to-1 Transcription Protocol (转录任务纯搬运原则：纯复制粘贴，严禁改动字号与格式)
**CRITICAL**: When transcribing (转录) from an existing draft OOS PDF (e.g. `OOS-XXXXXX ... .pdf`) into the official ZenQMS investigation template (e.g. `CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf`):
1. **Pure Copy-Paste Only (纯复制粘贴，不做任何格式加工)**:
   - Only copy `field_value` and checkbox states (`w.field_value`) 1-to-1 from the source document to the target document.
   - **DO NOT** attempt to calculate, override, or tamper with `text_fontsize` or text formatting.
   - Whatever text is in the source field, copy and paste it directly. Let the PDF viewer and official template render the fields naturally.
2. **Zero Unauthorized Formatting Modifications**:
   - The user has explicitly mandated: "就转录（复制粘贴）过去就行，不要管什么字体大小格式的".
   - Treat the field contents as pure data transfer without adding manual typography styling or font overrides during transcription.
"""

with open(r'.agents\AGENTS.md', 'a', encoding='utf-8') as f:
    f.write(section19)
print('Successfully appended Section 19 to AGENTS.md')
