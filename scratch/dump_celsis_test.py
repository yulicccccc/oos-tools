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
        
    tests_tab = page.locator("a:has-text('Tests')")
    if tests_tab.is_visible():
        tests_tab.click()
        page.wait_for_timeout(2000)
        
    celsis_link = page.locator("a:has-text('Celsis Sterility Test')")
    if celsis_link.is_visible():
        href = celsis_link.get_attribute("href")
        print("Celsis Test Href:", href)
        celsis_link.click()
        page.wait_for_timeout(3000)
        print("Celsis Test Details URL:", page.url)
        print("Celsis Test Details Title:", page.title())
        
        # Dump test details
        text = page.inner_text("body")
        with open("scratch/celsis_test_t1148992.txt", "w", encoding="utf-8") as f:
            f.write(text)
        page.screenshot(path="scratch/celsis_test_t1148992.png", full_page=True)
        
        # Dump all inputs/selects/values
        res = page.evaluate('''() => {
            const out = [];
            document.querySelectorAll('input, select, textarea').forEach(el => {
                const id = el.id || el.name;
                const val = el.tagName === 'SELECT' ? (el.options[el.selectedIndex]?.text || el.value) : el.value;
                if (val) out.push({id, val, tag: el.tagName});
            });
            return out;
        }''')
        for r in res:
            print(f"  {r['id']} = {r['val']}")
            
    browser.close()
