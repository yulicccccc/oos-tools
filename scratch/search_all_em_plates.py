import os, time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

samples_to_check = [
    'ETX-261005-0773',
    'ETX-260929-0335',
    'ETX-260929-0341',
    'ETX-260929-0344'
]

results = {}

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
    time.sleep(2)
    
    if '/Account/Login' in page.url:
        page.locator("input[name='Username']").fill('qchen')
        page.locator("input[type='submit']").click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
        time.sleep(2)
        
    for s_id in samples_to_check:
        print(f"\n==========================================")
        print(f"Searching for {s_id}...")
        
        # Select Submission #
        page.locator("#SearchFieldListDropdown").select_option("1")
        # Select Contains
        page.locator("#SearchMatchListDropdown").select_option("2")
        
        inp = page.locator("#srchCriteria")
        inp.fill(s_id)
        btn = page.locator("#FindButton")
        btn.click()
        time.sleep(4)
        
        page.screenshot(path=f'scratch/search_result_{s_id}.png', full_page=True)
        
        # Check table
        links = page.locator(f"a:has-text('{s_id}')")
        count = links.count()
        print(f"Found {count} links for {s_id}")
        
        if count > 0:
            sub_url = links.first.get_attribute("href")
            if sub_url.startswith("/"):
                sub_url = "https://etrax.eagleanalytical.com" + sub_url
            print(f"Navigating to {sub_url}...")
            
            p_detail = ctx.new_page()
            p_detail.goto(sub_url, wait_until='domcontentloaded')
            time.sleep(2)
            p_detail.screenshot(path=f'scratch/detail_{s_id}.png', full_page=True)
            
            # Click Tests tab
            t_tab = p_detail.locator("a[href='#Tests']").first
            if t_tab.count() > 0:
                t_tab.click()
                time.sleep(2)
                p_detail.screenshot(path=f'scratch/tests_{s_id}.png', full_page=True)
                
                # Look for test links
                t_links = p_detail.locator("div#Tests a[href*='SubmissionTest']").all()
                for idx, tl in enumerate(t_links):
                    t_href = tl.get_attribute("href")
                    if t_href.startswith("/"):
                        t_href = "https://etrax.eagleanalytical.com" + t_href
                    print(f"  Test {idx+1}: {t_href}")
                    
                    p_test = ctx.new_page()
                    p_test.goto(t_href, wait_until='domcontentloaded')
                    time.sleep(2)
                    
                    td_tab = p_test.locator("a[href='#TestDetails']").first
                    if td_tab.count() > 0:
                        td_tab.click()
                        time.sleep(2)
                    p_test.screenshot(path=f'scratch/test_details_{s_id}_{idx}.png', full_page=True)
                    
                    # Extract inputs and selects
                    content = p_test.content()
                    soup = BeautifulSoup(content, 'html.parser')
                    td_div = soup.find('div', id='TestDetails')
                    if td_div:
                        print(f"--- Test Details for {s_id} ---")
                        for form_grp in td_div.find_all('div', class_='form-group'):
                            lbl = form_grp.find('label')
                            lbl_txt = lbl.get_text(strip=True) if lbl else ''
                            inp_elem = form_grp.find(['input', 'select', 'textarea'])
                            val = ''
                            if inp_elem:
                                if inp_elem.name == 'select':
                                    opt = inp_elem.find('option', selected=True)
                                    val = opt.get_text(strip=True) if opt else ''
                                else:
                                    val = inp_elem.get('value', '')
                            if lbl_txt or val:
                                print(f"  {lbl_txt}: {val}")
                                
                    p_test.close()
            p_detail.close()
            
    ctx.close()
    print("\nSearch completed successfully!")
