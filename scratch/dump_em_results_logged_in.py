import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

test_urls = [
    ("ETX-260922-0347_Diff", "https://etrax.eagleanalytical.com/SubmissionTest/Details/71Xe8k1Q2aUGc8c2xFa9dw__"),
    ("ETX-260922-0347_ID", "https://etrax.eagleanalytical.com/SubmissionTest/Details/QCeWLJSF$FQGolgU0Doahg__"),
    ("ETX-260922-0349_ID", "https://etrax.eagleanalytical.com/SubmissionTest/Details/xiDTxtOQdhB9kDJFMkYuTg__"),
    ("ETX-260922-0356_Diff", "https://etrax.eagleanalytical.com/SubmissionTest/Details/dHrWth2yxbNMDB3obBUHrQ__"),
    ("ETX-260922-0356_ID", "https://etrax.eagleanalytical.com/SubmissionTest/Details/4nXZgYYBHHMrsW2zkuEdTg__")
]

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    # Login first
    page.goto("https://etrax.eagleanalytical.com/Account/Login", wait_until='domcontentloaded')
    time.sleep(2)
    if page.locator("input[name='Username']").count() > 0:
        print("Logging in as qchen...")
        page.fill("input[name='Username']", "qchen")
        page.click("input[type='submit']")
        time.sleep(4)

    for name, url in test_urls:
        print(f"\n--- {name}: {url} ---")
        page.goto(url, wait_until='networkidle')
        time.sleep(2)
        if '/Account/Login' in page.url:
            print("Redirected to login, re-logging in...")
            page.fill("input[name='Username']", "qchen")
            page.click("input[type='submit']")
            time.sleep(4)
            page.goto(url, wait_until='networkidle')
            time.sleep(2)
            
        td_tab = page.locator("a:has-text('Test Details')").first
        if td_tab.count() > 0:
            td_tab.click()
            time.sleep(1)
            
        page.screenshot(path=f"scratch/{name}.png", full_page=True)
        # Extract all text from main content
        txt = page.locator("#TestDetails, .container, body").first.inner_text()
        print(f"Text snippet for {name}:")
        for line in txt.split("\n"):
            line_str = line.strip()
            if any(k in line_str.lower() for k in ['result', 'gram', 'stain', 'microbial', 'status', 'organism', 'id', 'fail', 'pass', 'positive', 'rod', 'cocci', 'count', 'cfu', 'enterococcus', 'staphylococcus', 'bacillus', 'micrococcus']):
                print(f"    {line_str}")

    ctx.close()
