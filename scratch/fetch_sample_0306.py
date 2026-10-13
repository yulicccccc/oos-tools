import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
sample_url = 'https://etrax.eagleanalytical.com/Submission/Details/j$B8MDDTkRSGXhr3At81EQ__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    print(f"Navigating to sample URL: {sample_url}")
    page.goto(sample_url, wait_until='domcontentloaded')
    time.sleep(3)
    
    if "/Account/Login" in page.url:
        print("Redirected to login. Attempting auto-login...")
        inp = page.locator("input[name='Username'], input[id='Username']").first
        if inp.count() > 0:
            inp.fill('qchen')
            btn = page.locator("input[type='submit'], button[type='submit']").first
            btn.click()
            time.sleep(5)
            page.goto(sample_url, wait_until='domcontentloaded')
            time.sleep(3)
            
    print(f"Current page URL: {page.url}")
    print(f"Page Title: {page.title()}")
    
    page.screenshot(path='scratch/sample_0306_main.png', full_page=True)
    print("Saved scratch/sample_0306_main.png")
    
    content = page.content()
    with open('scratch/sample_0306_page.html', 'w', encoding='utf-8') as f:
        f.write(content)
        
    # Check for Tests tab
    tests_tab = page.locator("a:has-text('Tests')").first
    if tests_tab.count() > 0:
        print("Clicking Tests tab...")
        tests_tab.click()
        time.sleep(2)
        page.screenshot(path='scratch/sample_0306_tests.png', full_page=True)
        print("Saved scratch/sample_0306_tests.png")
        with open('scratch/sample_0306_tests.html', 'w', encoding='utf-8') as f:
            f.write(page.content())
            
    # Look for Scan RDI test links
    test_links = page.locator("a[href*='SubmissionTest'], a[href*='Test/Details'], a[href*='TestDetails']")
    print(f"Found {test_links.count()} test links:")
    scan_url = None
    for i in range(test_links.count()):
        tl = test_links.nth(i)
        txt = tl.inner_text().strip()
        href = tl.get_attribute("href")
        print(f"  Link {i+1}: '{txt}' -> {href}")
        if "Scan" in txt or "Sterility" in txt:
            scan_url = f"https://etrax.eagleanalytical.com{href}" if href.startswith("/") else href

    if scan_url:
        print(f"\nNavigating to Scan RDI Test Details: {scan_url}")
        page.goto(scan_url, wait_until='domcontentloaded')
        time.sleep(3)
        page.screenshot(path='scratch/sample_0306_scan_details.png', full_page=True)
        print("Saved scratch/sample_0306_scan_details.png")
        with open('scratch/sample_0306_scan_details.html', 'w', encoding='utf-8') as f:
            f.write(page.content())
            
    # Documents tab
    page.goto(sample_url, wait_until='domcontentloaded')
    time.sleep(2)
    doc_tab = page.locator("a:has-text('Documents')").first
    if doc_tab.count() > 0:
        print("Clicking Documents tab...")
        doc_tab.click()
        time.sleep(2)
        page.screenshot(path='scratch/sample_0306_docs.png', full_page=True)
        print("Saved scratch/sample_0306_docs.png")
        with open('scratch/sample_0306_docs.html', 'w', encoding='utf-8') as f:
            f.write(page.content())

    ctx.close()
