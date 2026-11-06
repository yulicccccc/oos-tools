import sys

sys.stdout.reconfigure(encoding='utf-8')

with open("PROJECT_STATE.md", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace("## Version: v12.4.2 (Previous: v12.4.1)", "## Version: v12.4.3 (Previous: v12.4.2)")

extra_note = """
- **Deep Forensic Table Grid Calibration & Zero Text Duplication (v12.4.3)**:
  - Discovered that previous programmatic table generation caused duplicate cell text (e.g. `None\\nNone`, `Date Before Testing\\nDate Before Testing`) due to merging cells after population.
  - Cloned the exact XML table grid from QA-approved benchmark `262403_PGRrx EM TABLE.docx`, cleaned body spacing with explicit page breaks, mapped exact unique cell column indices, and regenerated both DOCX and PDF deliverables.
  - Verified 100% elimination of duplicate text, flawless layout across all 3 table pages, active hyperlinks on P7 and P9, and zero overflow `+` signs across P1-6.
"""

if "### What Was Completed" in text:
    idx = text.find("### What Was Completed") + len("### What Was Completed")
    text = text[:idx] + extra_note + text[idx:]

with open("PROJECT_STATE.md", "w", encoding="utf-8") as f:
    f.write(text)

print("PROJECT_STATE.md updated to v12.4.3.")
