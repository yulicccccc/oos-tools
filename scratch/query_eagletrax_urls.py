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
            print(f"Navigated to Submission page: {p_sub.url}")
            
            # Get sample name, client, etc.
            client_elem = p_sub.locator(".client-name, #ClientName, a[href*='Client/Details']").first
            client_name = client_elem.text_content().strip() if client_elem.count() > 0 else "Unknown"
            
            sample_elem = p_sub.locator("#SampleName, input[name='SampleName']").first
            sample_name = sample_elem.input_value() if sample_elem.count() > 0 else "Unknown"
            if sample_name == "Unknown":
                sample_name = p_sub.locator("td:has-text('Sample Name') + td, label:has-text('Sample Name') + div").text_content() if p_sub.locator("td:has-text('Sample Name') + td").count() > 0 else "Unknown"
                
            print(f"Client: {client_name} | Sample: {sample_name}")
            
            # Check Tests tab
            t_tab = p_sub.locator("a[href='#Tests']").first
            if t_tab.count() > 0:
                t_tab.click()
                time.sleep(1)
                test_links = p_sub.locator("div#Tests a[href*='SubmissionTest/Details']").all()
                for tl in test_links:
                    t_href = tl.get_attribute("href")
                    t_url = f"https://etrax.eagleanalytical.com{t_href}" if t_href.startswith("/") else t_href
                    t_text = tl.text_content().strip()
                    print(f"Test Details Link: {t_url} (text: '{t_text}')")
            p_sub.close()
        else:
            print(f"No result found for {s_id}!")
            # Let's take screenshot of search result
            page.screenshot(path=f'scratch/no_result_{s_id}.png')
            
    ctx.close()
