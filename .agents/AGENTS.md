# Project-Scoped Rules for Eagle Analytical OOS Tools

## SOP: Creating a New OOS Module (Template Handoff Workflow)

Whenever the user initiates the creation of a new OOS module (e.g., adding a new test like USP71), follow this strict Human-AI collaborative SOP for the **Template Handoff (Step 1)**.

### The Contract
The core architecture of this system relies on `docxtpl` (Jinja2 syntax in Word). The Word document is the ultimate source of truth for all static text, formatting, and tables. Python code MUST NOT hardcode any long boilerplate text; it only populates variables marked with `{{variable_name}}`.

### Step 1 Workflow (Human & AI Responsibilities)

1. **Human Action**: 
   - The user duplicates an existing `.docx` template (e.g., `Celsis OOS P1 template 0.docx`) and renames it to the new test (e.g., `USP71 OOS P1 template 0.docx`).
   - The user manually updates the `{{}}` variable tags inside the Word document to match the specific needs of the new test (e.g., changing `{{celsis_result}}` to `{{usp71_id}}`).
   - The user **MUST CLOSE** Microsoft Word to release the file lock.
   - The user notifies the AI to proceed.

2. **AI Action**:
   - **Global Replace**: The AI runs a Python script (using `python-docx`) to globally search and replace the old test name (e.g., "Celsis") with the new test name (e.g., "USP <71>") across all paragraphs and tables in the new `.docx` file.
   - **Syntax 体检 (Sanity Check)**: The AI runs a Python script (using `re` and `python-docx`) to scan the document for malformed Jinja2 tags (e.g., spaces inside tags like `{{  subculture _name  }}`). The AI should automatically fix these syntax errors to prevent `docxtpl` from crashing.
   - **Variable Extraction (The Contract)**: Once the document is clean, the AI uses `docxtpl.DocxTemplate(filepath).get_undeclared_template_variables()` to extract the definitive list of all `{{}}` variables.
   - **Handoff**: The AI presents this alphabetized variable list to the user in the chat. This list acts as the absolute data contract for Step 2 (writing the `_logic.py` engine).

### Behavior Enforcement
- DO NOT attempt to modify a `.docx` file while the user has it open (it will throw `Permission denied`). Instruct the user to close the file first.
- NEVER skip the syntax check phase. MS Word formatting and typos often introduce invisible characters or spaces into `{{}}` tags which will fatally crash `docxtpl`.

---

## 🚨 ARCHITECTURAL CONSTRAINT: The Dual-Template Bulk Insertion

This architecture requires TWO parallel templates due to differing rendering constraints between Word and PDF:

### 1. `template.docx` (Standard Word Template)
- **Purpose**: The primary Word document for generation.
- **Rule**: It uses a single massive macro tag `{{ smart_phase1_summary }}` to ingest the entire narrative string from the backend logic in one go.

### 2. `template 0.docx` (PDF-Compatible Hollow Template)
- **Purpose**: The fallback created strictly as a workaround for `pypdf` limitations (PDF AcroForms have character limits on single fields).
- **Rule**: It splits the massive narrative into two giant placeholder macros: `{{ smart_phase1_part1 }}` and `{{ smart_phase1_part2 }}`.

### Python Backend Obligation
- **DO NOT** attempt to map granular English sentences inside the Word templates.
- **MUST** assemble complete, multi-paragraph text blocks (including domain-specific logic) entirely within the Python `_logic.py` engine.
- **MUST** push this massive variable block simultaneously to `smart_phase1_summary` (for Word) and `smart_phase1_part1/2` (for PDF).

---

## 🚨 Template Preservation & File History Rule

To prevent data loss and preserve user-created assets, you MUST follow these instructions:
1. **Never Overwrite Templates Directly**: Before making any modification or running scripts on template files (like `template 0.docx`, `template.docx`, `template.pdf`), you MUST create a copy of the original file in a `.history/` directory or rename it with a timestamp suffix (e.g. `_backup_YYYYMMDD_HHMMSS`) to preserve the historical version.
2. **Preserve User Assets**: Treat all user-made template documents as sacred. Never run bootstrap or tag-fixing scripts that overwrite them unless the user explicitly commands you to do so.

