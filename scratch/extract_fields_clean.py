import os
import sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

test_urls = [
    ('ETX-260922-0347 Diff', 'https://etrax.eagleanalytical.com/SubmissionTest/Details/71Xe8k1Q2aUGc8c2xFa9dw__'),
    ('ETX-260922-0347 ID', 'https://etrax.eagleanalytical.com/SubmissionTest/Details/QCeWLJSF$FQGolgU0Doahg__'),
    ('ETX-260922-0349 ID', 'https://etrax.eagleanalytical.com/SubmissionTest/Details/xiDTxtOQdhB9kDJFMkYuTg__'),
    ('ETX-260922-0356 Diff', 'https://etrax.eagleanalytical.com/SubmissionTest/Details/dHrWth2yxbNMDB3obBUHrQ__'),
    ('ETX-260922-0356 ID', 'https://etrax.eagleanalytical.com/SubmissionTest/Details/4nXZgYYBHHMrsW2zkuEdTg__')
]

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    for name, url in test_urls:
        page.goto(url, wait_until='networkidle')
        js = """() => {
            const results = [];
            document.querySelectorAll('input, select, textarea').forEach(el => {
                let val = el.value;
                if (el.tagName.toLowerCase() === 'select') {
                    val = el.options[el.selectedIndex]?.text || '';
                }
                if (val && val.trim() && val !== '- Select -' && !val.includes('__RequestVerificationToken')) {
                    const id = el.id || el.name || '';
                    results.push({id: id, val: val.trim()});
                }
            });
            return results;
        }"""
        data = page.evaluate(js)
        print(f"=== {name} ===")
        for item in data:
            if not any(k in item['id'].lower() for k in ['token', 'button', 'filter']):
                print(f"  {item['id']}: {item['val']}")
    ctx.close()
