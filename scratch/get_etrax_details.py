import os
import time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
sample_url = 'https://etrax.eagleanalytical.com/Submission/Details/08VuIoFsBaXGehnOxJ3PMQ__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    page.goto(sample_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    # Auto login if needed
    if "/Account/Login" in page.url or "Login" in page.title():
        print("Logging in with qchen...")
        inp = page.locator("input[name='Username'], input[id='Username']").first
        if inp.count() > 0:
            inp.fill('qchen')
            btn = page.locator("input[type='submit'], button[type='submit']").first
            btn.click()
            time.sleep(5)
            page.goto(sample_url, wait_until='domcontentloaded')
            time.sleep(3)
            
    print(f"Current page: {page.url} ({page.title()})")
    
    # Click Tests tab
    tab = page.locator("a[href='#Tests']").first
    print("Clicking a[href='#Tests']...")
    tab.click()
    time.sleep(3)
    
    page.screenshot(path='scratch/tests_tab_loaded.png', full_page=True)
    print("Saved scratch/tests_tab_loaded.png")
    
    # Save Tests tab content
    content = page.content()
    with open('scratch/tests_tab_page.html', 'w', encoding='utf-8') as f:
        f.write(content)
        
    soup = BeautifulSoup(content, 'html.parser')
    tests_div = soup.find('div', id='Tests')
    if tests_div:
        print("\n--- Rows in div#Tests ---")
        for tr in tests_div.find_all('tr'):
            cells = [c.get_text(strip=True) for c in tr.find_all(['th', 'td'])]
            links = [a.get('href') for a in tr.find_all('a')]
            print("Row:", cells, "Links:", links)
            
    # Find any SubmissionTest link
    test_links = page.locator("div#Tests a[href*='SubmissionTest'], div#Tests a[href*='Test/Details']")
    count = test_links.count()
    print(f"\nFound {count} links inside div#Tests")
    if count > 0:
        target_url = test_links.first.get_attribute("href")
        if target_url.startswith("/"):
            target_url = "https://etrax.eagleanalytical.com" + target_url
        print(f"Navigating to Test page: {target_url}")
        page.goto(target_url, wait_until='domcontentloaded')
        time.sleep(3)
        page.screenshot(path='scratch/test_subpage.png', full_page=True)
        print("Saved scratch/test_subpage.png")
        
        # Click Test Details tab
        td_tab = page.locator("a[href='#TestDetails'], a:has-text('Test Details')").first
        if td_tab.count() > 0:
            print("Clicking Test Details tab...")
            td_tab.click()
            time.sleep(3)
            page.screenshot(path='scratch/test_details_tab.png', full_page=True)
            print("Saved scratch/test_details_tab.png")
            
        with open('scratch/test_subpage.html', 'w', encoding='utf-8') as f:
            f.write(page.content())
            
    ctx.close()
