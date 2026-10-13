import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
sample_url = 'https://etrax.eagleanalytical.com/Submission/Details/08VuIoFsBaXGehnOxJ3PMQ__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    page.goto(sample_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    # Click Tests tab
    tests_tab = page.locator("a:has-text('Tests')").first
    print("Clicking Tests tab...")
    tests_tab.click()
    time.sleep(2)
    page.screenshot(path='scratch/sample_0148_tests_tab.png', full_page=True)
    print("Saved scratch/sample_0148_tests_tab.png")
    
    # Look for links in the Tests table
    links = page.locator("a")
    for i in range(links.count()):
        l = links.nth(i)
        href = l.get_attribute("href") or ""
        txt = l.inner_text().strip()
        if "SubmissionTest" in href or "Test" in href or "Scan" in txt:
            print(f"Test link: '{txt}' -> {href}")
            
    # Click the first test link (or details link)
    target_link = page.locator("a[href*='SubmissionTest'], a[href*='Test/Details']").first
    if target_link.count() > 0:
        print(f"Navigating to test link: {target_link.get_attribute('href')}")
        target_link.click()
        time.sleep(3)
        page.screenshot(path='scratch/sample_0148_test_page.png', full_page=True)
        print("Saved scratch/sample_0148_test_page.png")
        
        # Click Test Details tab if present
        td_tab = page.locator("a:has-text('Test Details')").first
        if td_tab.count() > 0:
            print("Clicking Test Details tab...")
            td_tab.click()
            time.sleep(2)
            page.screenshot(path='scratch/sample_0148_test_details.png', full_page=True)
            print("Saved scratch/sample_0148_test_details.png")
            
        with open('scratch/test_details_page.html', 'w', encoding='utf-8') as f:
            f.write(page.content())
            
    ctx.close()
