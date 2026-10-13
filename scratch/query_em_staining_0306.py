import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

samples = [
    'ETX-260922-0347',
    'ETX-260922-0349',
    'ETX-260922-0356'
]

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='networkidle')
    
    if '/Account/Login' in page.url or 'Login' in page.title():
        page.locator("input[name='Username']").fill('qchen')
        page.locator("input[type='submit']").click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='networkidle')
        
    for s_id in samples:
        print(f"\n==========================================")
        print(f"Querying EagleTrax for {s_id}...")
        
        page.locator("#SearchFieldListDropdown").select_option("1")
        page.locator("#SearchMatchListDropdown").select_option("2")
        page.locator("#srchCriteria").fill(s_id)
        page.locator("#FindButton").click()
        time.sleep(3)
        
        links = page.locator(f"a:has-text('{s_id}')")
        cnt = links.count()
        print(f"Found {cnt} link(s) matching '{s_id}'")
        
        if cnt > 0:
            href = links.first.get_attribute("href")
            sub_url = f"https://etrax.eagleanalytical.com{href}" if href.startswith("/") else href
            print(f"Submission URL: {sub_url}")
            
            p_sub = ctx.new_page()
            p_sub.goto(sub_url, wait_until='networkidle')
            time.sleep(2)
            
            # Print sample info
            sample_name = p_sub.locator("input#SampleName, #SampleName").first
            s_name_val = sample_name.input_value() if sample_name.count() > 0 else "N/A"
            print(f"Sample Name: {s_name_val}")
            
            # Tests tab
            t_tab = p_sub.locator("a:has-text('Tests')").first
            if t_tab.count() > 0:
                t_tab.click()
                time.sleep(2)
                t_links = p_sub.locator("a[href*='SubmissionTest'], a[href*='Test/Details']")
                print(f"Tests found: {t_links.count()}")
                for i in range(t_links.count()):
                    tl = t_links.nth(i)
                    print(f"  Test {i+1}: '{tl.inner_text().strip()}' -> {tl.get_attribute('href')}")
                    
                for i in range(t_links.count()):
                    tl = t_links.nth(i)
                    t_txt = tl.inner_text().strip()
                    t_href = tl.get_attribute("href")
                    full_t_url = f"https://etrax.eagleanalytical.com{t_href}" if t_href.startswith("/") else t_href
                    p_test = ctx.new_page()
                    p_test.goto(full_t_url, wait_until='networkidle')
                    time.sleep(2)
                    
                    # Click Test Details tab if available
                    td_tab = p_test.locator("a:has-text('Test Details')").first
                    if td_tab.count() > 0:
                        td_tab.click()
                        time.sleep(1)
                        
                    # Extract test results text
                    body_txt = p_test.locator("#TestDetails, .test-details, body").first.inner_text()
                    lines = [l.strip() for l in body_txt.split('\n') if l.strip()]
                    print(f"\n  --- Details for {t_txt} ({s_id}) ---")
                    for l in lines[:25]:
                        if any(k in l.lower() for k in ['result', 'gram', 'stain', 'microbial', 'status', 'organism', 'id', 'fail', 'pass', 'positive', 'rod', 'cocci', 'count']):
                            print(f"    {l}")
                    p_test.close()
            p_sub.close()

    ctx.close()
