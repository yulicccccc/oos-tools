import os
from playwright.sync_api import sync_playwright

user_data_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
url_sample = "https://etrax.eagleanalytical.com/Submission/Details/fZlFl0vWDpqR-VWBsez7IA__"

with sync_playwright() as p:
    browser = p.chromium.launch_persistent_context(user_data_dir=user_data_dir, headless=True)
    page = browser.new_page()
    page.goto(url_sample, wait_until='domcontentloaded')
    page.wait_for_timeout(3000)
    
    if "Log in" in page.title():
        page.fill("input[name='Username']", "qchen")
        page.click("input[type='submit'], button[type='submit']")
        page.wait_for_timeout(3000)
        
    print("Page URL:", page.url)
    
    tabs = page.locator("ul.nav-tabs li a, .nav-pills li a").all()
    print(f"Found {len(tabs)} tabs:")
    for t in tabs:
        print("  Tab:", t.inner_text())
        
    # Click Tests tab
    tests_tab = page.locator("a:has-text('Tests')")
    if tests_tab.is_visible():
        tests_tab.click()
        page.wait_for_timeout(2000)
        print("\n--- TESTS TAB ---")
        print(page.locator("#Tests, .tab-pane.active").inner_text())
        page.screenshot(path="scratch/sub_tests_tab.png")
        
    # Click Event History tab
    eh_tab = page.locator("a:has-text('Event History')")
    if eh_tab.is_visible():
        eh_tab.click()
        page.wait_for_timeout(2000)
        print("\n--- EVENT HISTORY TAB ---")
        print(page.locator("#EventHistory, .tab-pane.active").inner_text()[:1000])
        
    browser.close()
