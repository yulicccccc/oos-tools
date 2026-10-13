import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

test_urls = [
    ("ETX-260922-0347 Diff", "https://etrax.eagleanalytical.com/SubmissionTest/Details/71Xe8k1Q2aUGc8c2xFa9dw__"),
    ("ETX-260922-0347 ID", "https://etrax.eagleanalytical.com/SubmissionTest/Details/QCeWLJSF$FQGolgU0Doahg__"),
    ("ETX-260922-0349 ID", "https://etrax.eagleanalytical.com/SubmissionTest/Details/xiDTxtOQdhB9kDJFMkYuTg__"),
    ("ETX-260922-0356 Diff", "https://etrax.eagleanalytical.com/SubmissionTest/Details/dHrWth2yxbNMDB3obBUHrQ__"),
    ("ETX-260922-0356 ID", "https://etrax.eagleanalytical.com/SubmissionTest/Details/4nXZgYYBHHMrsW2zkuEdTg__")
]

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    for name, url in test_urls:
        print(f"\n--- {name}: {url} ---")
        page.goto(url, wait_until='networkidle')
        time.sleep(2)
        # Click Test Details tab if present
        td_tab = page.locator("a:has-text('Test Details')").first
        if td_tab.count() > 0:
            td_tab.click()
            time.sleep(1)
            
        page.screenshot(path=f"scratch/{name.replace(' ', '_')}.png")
        # Dump inputs, selects, textareas
        inputs = page.locator("input, select, textarea")
        for i in range(inputs.count()):
            elem = inputs.nth(i)
            tag = elem.evaluate("el => el.tagName.toLowerCase()")
            val = elem.input_value() if tag in ['input', 'textarea'] else elem.evaluate("el => el.options[el.selectedIndex]?.text || ''")
            label = elem.evaluate("""el => {
                const id = el.id;
                if (id) {
                    const l = document.querySelector(`label[for='${id}']`);
                    if (l) return l.innerText;
                }
                const parent = el.closest('.form-group, tr');
                return parent ? parent.innerText.split('\\n')[0] : '';
            }""")
            if val and val.strip() and val != '- Select -':
                print(f"  [{label.strip()}]: {val.strip()}")

    ctx.close()
