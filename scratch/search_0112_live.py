import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
profile_dir = os.path.expandvars(r'%LOCALAPPDATA%\pastdue_playwright_session')

with sync_playwright() as p:
    browser = p.chromium.launch_persistent_context(
        user_data_dir=profile_dir,
        headless=True
    )
    page = browser.new_page()
    
    print("Navigating to Submission...")
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='networkidle')
    print("Page URL:", page.url)
    print("Page Title:", page.title())
    
    # Click Tests tab if available
    tests_tab = page.locator("a:has-text('Tests')").first
    if tests_tab.count() > 0:
        print("Clicking Tests tab...")
        tests_tab.click()
        time.sleep(2)
        print("URL after clicking Tests:", page.url)
    
    # Find search input
    search_inp = page.locator("input[placeholder*='Search'], input[type='search'], input[name*='search'], input[id*='search']").first
    if search_inp.count() > 0:
        print("Filling search with ETX-260901-0112...")
        search_inp.fill('ETX-260901-0112')
        search_inp.press('Enter')
        time.sleep(3)
        
        # Check rows
        rows = page.locator("tr").all()
        print(f"Rows count: {len(rows)}")
        for r in rows:
            txt = r.inner_text().strip().replace('\n', ' | ')
            if '260901-0112' in txt or '0112' in txt:
                print("Found matching row:", txt)
                links = r.locator("a").all()
                for l in links:
                    print("  Link:", l.get_attribute('href'), l.inner_text())
    
    page.screenshot(path='scratch/search_0112_tests_result.png')
    browser.close()
