import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

etx_ids = [
    "ETX-260922-0347",
    "ETX-260922-0349",
    "ETX-260922-0350"
]

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    # Go to Submissions list
    page.goto("https://etrax.eagleanalytical.com/Submission", wait_until='domcontentloaded')
    time.sleep(2)
    if '/Account/Login' in page.url:
        page.fill('input[name="Username"]', 'qchen')
        page.click('input[type="submit"]')
        time.sleep(5)
        page.goto("https://etrax.eagleanalytical.com/Submission", wait_until='domcontentloaded')
        time.sleep(2)
        
    for etx_id in etx_ids:
        print(f"\nSearching for {etx_id}...")
        # The main submission list has search inputs
        # Let's inspect inputs
        inp = page.locator("input#SearchString, input[name='SearchString'], input[id*='Search']").first
        if inp.count() == 0:
            # Maybe generic input
            inp = page.locator("input.form-control").first
        inp.fill(etx_id)
        # click Find
        btn = page.locator("input[value='Find'], button:has-text('Find')").first
        btn.click()
        time.sleep(3)
        
        # Take screenshot of search results
        page.screenshot(path=f"scratch/search_{etx_id}.png")
        
        link = page.locator(f"a:has-text('{etx_id}')").first
        if link.count() > 0:
            href = link.get_attribute("href")
            full_sub_url = f"https://etrax.eagleanalytical.com{href}" if href.startswith("/") else href
            print(f"Found {etx_id} URL: {full_sub_url}")
            
            # Go to submission
            page.goto(full_sub_url, wait_until='domcontentloaded')
            time.sleep(2)
            page.screenshot(path=f"scratch/page_{etx_id}.png")
            
            # Go to Tests tab
            t_tab = page.locator("a:has-text('Tests')").first
            if t_tab.count() > 0:
                t_tab.click()
                time.sleep(2)
                page.screenshot(path=f"scratch/tests_tab_{etx_id}.png")
                
                # Check for test links
                t_links = page.locator("a[href*='SubmissionTest'], a[href*='Test/Details']")
                for i in range(t_links.count()):
                    tl = t_links.nth(i)
                    print(f"  Test: '{tl.inner_text().strip()}' -> {tl.get_attribute('href')}")
                    
                # If there's a test link, open the first one
                if t_links.count() > 0:
                    t_url = t_links.first.get_attribute("href")
                    if t_url:
                        full_t_url = f"https://etrax.eagleanalytical.com{t_url}" if t_url.startswith("/") else t_url
                        page.goto(full_t_url, wait_until='domcontentloaded')
                        time.sleep(2)
                        page.screenshot(path=f"scratch/test_detail_{etx_id}.png")
                        with open(f"scratch/test_detail_{etx_id}.html", "w", encoding="utf-8") as f:
                            f.write(page.content())
            
            # Back to submission search
            page.goto("https://etrax.eagleanalytical.com/Submission", wait_until='domcontentloaded')
            time.sleep(2)
        else:
            print(f"Link not found for {etx_id}")

    ctx.close()
