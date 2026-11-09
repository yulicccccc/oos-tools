with open('PROJECT_STATE.md', 'r', encoding='utf-8') as f:
    text = f.read()

new_section = '''## Version: v12.4.25 (Previous: v12.4.24)

## Current Phase: Page 3 Section D Narrative Streamlining (Elimination of Redundant Processing Sequence) (v12.4.25)

### What Was Completed
- **Streamlined Page 3 Section D Testing Narrative (`Text Field49`)**:
  - Eliminated redundant sequence phrasing `(12th sample processed, hs12)` from Paragraph 8:
    - Revised to: *"On 20Oct26, a rapid sterility test was performed on the sample using the ScanRDI method under test record 102026-2606-1 (Analysis ID 102026-2606-1034). The sample was prepped by Nicholas Trinh, processed by Halaina Smith, changeover performed by America Alanis, and subsequently read by Maryam Naeem. The test revealed 11 total events and 1 confirmed microbial event exhibiting rod shaped morphology, see Table 1."*
    - Retained sequence defense (`processed as the 12th sample by Halaina Smith`) on Page 4 (`Text Field50`) where sequence comparison vs 53rd sample (`OOS-262395`) serves as critical cross-contamination evidence.
- **Verification & Zero-Overflow Audit**:
  - Zero-Overflow Verification: **0 overflowing fields detected** across all form pages.
  - Page count: Strictly 7 pages (6 pages ZenQMS AcroForm + 1 page Table 1/2 standalone page).
  - Synchronized versioned copy to Desktop: `CORP-FORM-21- 20 Oct 2026 OOS-262394 Optio RX (E18631) - ScanRDI (v12.4.25).pdf`.

### Previous Releases
- **Page 2 Section B Equipment & Calibration Table Alignment (10-Item Roster) (v12.4.24)**'''

split_point = text.find('### Previous Releases')
rest = text[split_point:]

header = '# Project State: OOS Tools\n\n'
new_text = header + new_section + '\n' + rest[len('### Previous Releases'):]

with open('PROJECT_STATE.md', 'w', encoding='utf-8') as f:
    f.write(new_text)

print('Updated PROJECT_STATE.md to v12.4.25!')
