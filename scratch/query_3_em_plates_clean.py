import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

samples = [
    'ETX-260922-0347',
    'ETX-260922-0349',
    'ETX-260922-0350'
]

results = {}

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='networkidle')
    
    if '/Account/Login' in page.url or 'Login' in page.title():
        print("Logging in...")
        page.locator("input[name='Username']").fill('qchen')
        page.locator("input[type='submit']").click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='networkidle')
        
    print(f"Logged in successfully. Current URL: {page.url}")
    
    for s_id in samples:
        print(f"\n==========================================")
        print(f"Querying EagleTrax for {s_id}...")
        
        # Select Submission #
        page.locator("#SearchFieldListDropdown").select_option("1")
        # Select Contains
        page.locator("#SearchMatchListDropdown").select_option("2")
        
        inp = page.locator("#srchCriteria")
        inp.fill(s_id)
        btn = page.locator("#FindButton")
        btn.click()
        time.sleep(3)
        
        # Find links
        links = page.locator(f"a:has-text('{s_id}')")
        cnt = links.count()
        print(f"Found {cnt} link(s) matching '{s_id}'")
        
        if cnt > 0:
            href = links.first.get_attribute("href")
            full_sub_url = f"https://etrax.eagleanalytical.com{href}" if href.startswith("/") else href
            print(f"Submission URL: {full_sub_url}")
            
            p_sub = ctx.new_page()
            p_sub.goto(full_sub_url, wait_until='networkidle')
            time.sleep(2)
            
            # Go to Tests tab
            t_tab = p_sub.locator("a:has-text('Tests')").first
            if t_tab.count() > 0:
                t_tab.click()
                time.sleep(2)
                p_sub.screenshot(path=f"scratch/tests_{s_id}.png")
                
                # Check for test links
                test_links = p_sub.locator("a[href*='SubmissionTest'], a[href*='Test/Details']")
                print(f"Found {test_links.count()} test links for {s_id}")
                for i in range(test_links.count()):
                    tl = test_links.nth(i)
                    print(f"  Test {i+1}: '{tl.inner_text().strip()}' -> {tl.get_attribute('href')}")
                    
                if test_links.count() > 0:
                    t_url = test_links.first.get_attribute("href")
                    full_t_url = f"https://etrax.eagleanalytical.com{t_url}" if t_url.startswith("/") else t_url
                    p_sub.goto(full_t_url, wait_until='networkidle')
                    time.sleep(2)
                    p_sub.screenshot(path=f"scratch/test_details_{s_id}.png")
                    with open(f"scratch/test_details_{s_id}.html", "w", encoding="utf-8") as f:
                        f.write(p_sub.content())
            
            p_sub.close()
            results[s_id] = full_sub_url
        else:
            print(f"No link found for {s_id}")
            page.screenshot(path=f"scratch/search_fail_{s_id}.png")

    ctx.close()
    print("\nResults summary:")
    for k, v in results.items():
        print(f"  {k} -> {v}")
