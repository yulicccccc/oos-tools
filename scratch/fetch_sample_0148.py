import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
sample_url = 'https://etrax.eagleanalytical.com/Submission/Details/08VuIoFsBaXGehnOxJ3PMQ__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    print(f"Navigating to sample URL: {sample_url}")
    page.goto(sample_url, wait_until='domcontentloaded')
    time.sleep(3)
    
    # Check if redirect to login
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
    
    page.screenshot(path='scratch/sample_0148_main.png', full_page=True)
    print("Saved scratch/sample_0148_main.png")
    
    # Extract page text/tables
    content = page.content()
    with open('scratch/sample_0148_page.html', 'w', encoding='utf-8') as f:
        f.write(content)
        
    # Check for Test links or Test Details links
    test_links = page.locator("a[href*='SubmissionTest'], a[href*='TestDetails'], a[href*='Test/Details']")
    print(f"Found {test_links.count()} test links:")
    for i in range(test_links.count()):
        tl = test_links.nth(i)
        txt = tl.inner_text().strip()
        href = tl.get_attribute("href")
        print(f"  Link {i+1}: '{txt}' -> {href}")
        
    # If there are links to tests, let's navigate to the Scan RDI test
    for i in range(test_links.count()):
        tl = test_links.nth(i)
        txt = tl.inner_text().strip()
        href = tl.get_attribute("href")
        if "Scan" in txt or "Sterility" in txt or "SubmissionTest" in (href or ""):
            full_test_url = f"https://etrax.eagleanalytical.com{href}" if href.startswith("/") else href
            print(f"\nNavigating to Test: {full_test_url}")
            page.goto(full_test_url, wait_until='domcontentloaded')
            time.sleep(3)
            page.screenshot(path='scratch/sample_0148_test.png', full_page=True)
            print("Saved scratch/sample_0148_test.png")
            
            # Click Test Details tab if present
            td_tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
            if td_tab.count() > 0:
                print("Clicking Test Details tab...")
                td_tab.first.click()
                time.sleep(2)
                page.screenshot(path='scratch/sample_0148_test_details.png', full_page=True)
                print("Saved scratch/sample_0148_test_details.png")
                
            with open('scratch/sample_0148_test.html', 'w', encoding='utf-8') as f:
                f.write(page.content())
            break
            
    ctx.close()
    print("Done!")
