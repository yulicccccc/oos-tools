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

results = {}

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    for etx_id in etx_ids:
        print(f"\nSearching for {etx_id}...")
        page.goto("https://etrax.eagleanalytical.com/Submission/Search", wait_until='domcontentloaded')
        time.sleep(2)
        if '/Account/Login' in page.url:
            page.fill('input[name="Username"]', 'qchen')
            page.click('input[type="submit"]')
            time.sleep(5)
            page.goto("https://etrax.eagleanalytical.com/Submission/Search", wait_until='domcontentloaded')
            time.sleep(2)
            
        inp = page.locator("input[name='SubmissionNumber'], input[id='SubmissionNumber']").first
        if inp.count() == 0:
            # Maybe search input
            inp = page.locator("input[type='text']").first
        inp.fill(etx_id)
        btn = page.locator("input[value='Find'], button:has-text('Find')").first
        btn.click()
        time.sleep(3)
        
        # Check result
        sub_link = page.locator(f"a:has-text('{etx_id}')").first
        if sub_link.count() > 0:
            href = sub_link.get_attribute("href")
            sub_url = f"https://etrax.eagleanalytical.com{href}" if href.startswith("/") else href
            print(f"Found submission URL: {sub_url}")
            
            # Go to submission
            page.goto(sub_url, wait_until='domcontentloaded')
            time.sleep(2)
            page.screenshot(path=f"scratch/sub_{etx_id}.png", full_page=True)
            
            # Click Tests tab
            tests_tab = page.locator("a:has-text('Tests')").first
            if tests_tab.count() > 0:
                tests_tab.click()
                time.sleep(2)
                page.screenshot(path=f"scratch/tests_{etx_id}.png", full_page=True)
                
            # Look for test links
            test_links = page.locator("a[href*='SubmissionTest'], a[href*='Test/Details'], a[href*='TestDetails']")
            test_url = None
            for i in range(test_links.count()):
                tl = test_links.nth(i)
                txt = tl.inner_text().strip()
                t_href = tl.get_attribute("href")
                print(f"  Test: '{txt}' -> {t_href}")
                if "Differential" in txt or "Gram" in txt or "Microbial" in txt or "Plate" in txt:
                    test_url = f"https://etrax.eagleanalytical.com{t_href}" if t_href.startswith("/") else t_href
            
            # If no specific name, just click the first test link
            if not test_url and test_links.count() > 0:
                t_href = test_links.first.get_attribute("href")
                test_url = f"https://etrax.eagleanalytical.com{t_href}" if t_href.startswith("/") else t_href
                
            if test_url:
                page.goto(test_url, wait_until='domcontentloaded')
                time.sleep(2)
                page.screenshot(path=f"scratch/test_details_{etx_id}.png", full_page=True)
                
                # Click Test Details tab if exists
                td_tab = page.locator("a:has-text('Test Details')").first
                if td_tab.count() > 0:
                    td_tab.click()
                    time.sleep(2)
                    page.screenshot(path=f"scratch/test_details_tab_{etx_id}.png", full_page=True)
                
                # Grab text
                content_txt = page.content()
                with open(f"scratch/test_details_{etx_id}.html", "w", encoding="utf-8") as f:
                    f.write(content_txt)
                    
            results[etx_id] = {
                "sub_url": sub_url,
                "test_url": test_url
            }
        else:
            print(f"No link found for {etx_id}!")
            page.screenshot(path=f"scratch/no_link_{etx_id}.png")

    ctx.close()
    print("Done checking EM plates.")
