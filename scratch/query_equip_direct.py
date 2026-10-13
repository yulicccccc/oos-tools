import os, time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Equipment', wait_until='domcontentloaded')
    time.sleep(2)
    if '/Account/Login' in page.url:
        page.locator("input[name='Username']").fill('qchen')
        page.locator("input[type='submit']").click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Equipment', wait_until='domcontentloaded')
        time.sleep(2)
        
    print("Calling RefreshEagleEquipmentList for 1937...")
    page.evaluate("RefreshEagleEquipmentList(1, '', '1937')")
    time.sleep(3)
    
    html = page.locator("#EagleEquipmentList").inner_html()
    soup = BeautifulSoup(html, 'html.parser')
    for tr in soup.find_all('tr'):
        cells = [td.get_text(strip=True) for td in tr.find_all(['th', 'td'])]
        links = [a.get('href') for a in tr.find_all('a')]
        print("1937 Row:", cells, links)
        
    print("\nCalling RefreshEagleEquipmentList for 1312...")
    page.evaluate("RefreshEagleEquipmentList(1, '', '1312')")
    time.sleep(3)
    
    html = page.locator("#EagleEquipmentList").inner_html()
    soup = BeautifulSoup(html, 'html.parser')
    for tr in soup.find_all('tr'):
        cells = [td.get_text(strip=True) for td in tr.find_all(['th', 'td'])]
        links = [a.get('href') for a in tr.find_all('a')]
        print("1312 Row:", cells, links)
        
    ctx.close()
