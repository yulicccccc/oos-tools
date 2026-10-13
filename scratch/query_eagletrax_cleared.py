import os, time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

samples = [
    'ETX-260921-0148',
    'ETX-261005-0773',
    'ETX-260929-0335',
    'ETX-260929-0341',
    'ETX-260929-0344'
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
        print(f"Searching for {s_id}...")
        
        # Click Clear button first!
        page.locator("button:has-text('Clear'), a:has-text('Clear'), input[value='Clear'], #ClearButton").first.click()
        time.sleep(2)
        
        # Click 'Submissions' tab
        sub_tab = page.locator("a:has-text('Submissions')").first
        if sub_tab.count() > 0:
            sub_tab.click()
            time.sleep(1)
            
        # Select Submission #
        page.locator("#SearchFieldListDropdown").select_option("1")
        # Select Contains
        page.locator("#SearchMatchListDropdown").select_option("2")
        
        inp = page.locator("#srchCriteria")
        inp.fill(s_id)
        btn = page.locator("#FindButton")
        btn.click()
        time.sleep(3)
        
        page.screenshot(path=f'scratch/cleared_search_{s_id}.png')
        
        # Find links
        links = page.locator(f"a:has-text('{s_id}')")
        cnt = links.count()
        print(f"Found {cnt} link(s) matching '{s_id}'")
        for i in range(cnt):
            href = links.nth(i).get_attribute("href")
            txt = links.nth(i).text_content().strip()
            print(f"  Link {i+1}: text='{txt}', href='{href}'")
            
    ctx.close()
