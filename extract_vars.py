import sys
from docxtpl import DocxTemplate

def extract_vars(filepath):
    try:
        tpl = DocxTemplate(filepath)
        vars = tpl.get_undeclared_template_variables()
        print("\n" + "="*40)
        print("FOUND THE FOLLOWING {{VARIABLES}}:")
        print("="*40)
        for v in sorted(vars):
            print(f" - {v}")
        print("="*40)
    except Exception as e:
        print(f"[-] Error: {e}")

if __name__ == "__main__":
    extract_vars('USP71 OOS P1 template.docx')
