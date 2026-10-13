import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

samples_to_check = ['ETX-260929-0335', 'ETX-260929-0341', 'ETX-260929-0344']

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    # Go to Find Submissions
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
    time.sleep(2)
    
    if "/Account/Login" in page.url or "Login" in page.title():
        inp = page.locator("input[name='Username']").first
        inp.fill('qchen')
        page.locator("input[type='submit']").first.click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
        time.sleep(2)
        
    for s_id in samples_to_check:
        print(f"\nSearching for {s_id}...")
        search_inp = page.locator("input[type='search'], input[id*='Search'], #SubmissionSearch, input.form-control").first
        if search_inp.count() > 0:
            search_inp.fill(s_id)
            search_inp.press("Enter")
            time.sleep(3)
            
            # Check results
            links = page.locator(f"a:has-text('{s_id}')")
            if links.count() > 0:
                print(f"Found link for {s_id}: {links.first.get_attribute('href')}")
                detail_url = "https://etrax.eagleanalytical.com" + links.first.get_attribute('href')
                
                # Open in new page or navigate
                p_sub = ctx.new_page()
                p_sub.goto(detail_url, wait_until='domcontentloaded')
                time.sleep(2)
                p_sub.screenshot(path=f'scratch/em_{s_id}.png', full_page=True)
                
                # Check Tests tab
                t_tab = p_sub.locator("a[href='#Tests']").first
                if t_tab.count() > 0:
                    t_tab.click()
                    time.sleep(2)
                    test_link = p_sub.locator("div#Tests a[href*='SubmissionTest']").first
                    if test_link.count() > 0:
                        p_sub.goto("https://etrax.eagleanalytical.com" + test_link.get_attribute('href'), wait_until='domcontentloaded')
                        time.sleep(2)
                        td_tab = p_sub.locator("a[href='#TestDetails']").first
                        if td_tab.count() > 0:
                            td_tab.click()
                            time.sleep(1)
                        p_sub.screenshot(path=f'scratch/em_test_{s_id}.png', full_page=True)
                        print(f"Captured test details for {s_id}")
                p_sub.close()
            else:
                print(f"No link found for {s_id}")
                
    ctx.close()
