import os
import time
import re
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
print(f"Using Session Directory: {session_dir}")

target_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/qnwcLQO5BWeJhWCBG7jl8Q__#TestDetails'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        session_dir,
        headless=True  # try headless first with saved cookies
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    print(f"Navigating to: {target_url}")
    page.goto(target_url, wait_until='domcontentloaded')
    time.sleep(3)
    
    print(f"Current URL: {page.url}")
    page.screenshot(path="scratch/etrax_0520.png")
    
    # Check if login required
    if "Login" in page.url or "microsoft" in page.url:
        print("Login required! Re-launching in headed mode...")
        ctx.close()
        exit(2)
        
    # Click Test Details tab if available
    tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
    if tab.count() > 0:
        try:
            tab.first.click()
            time.sleep(2)
        except Exception as e:
            print("Tab click error:", e)
            
    page.screenshot(path="scratch/etrax_0520_details.png")
    
    # Extract all text on the page
    body_text = page.locator("body").inner_text()
    with open("scratch/etrax_0520_text.txt", "w", encoding="utf-8") as f:
        f.write(body_text)
    print("Saved body text to scratch/etrax_0520_text.txt")
    
    # Print selects
    selects = page.locator("select")
    for i in range(selects.count()):
        sel = selects.nth(i)
        sel_id = sel.get_attribute("id") or sel.get_attribute("name")
        val = page.evaluate("(el) => el.options[el.selectedIndex] ? el.options[el.selectedIndex].text : ''", sel.element_handle())
        print(f"Select [{sel_id}] = '{val}'")
        
    # Print inputs
    inputs = page.locator("input[type='text'], input[id*='Result'], textarea")
    for i in range(inputs.count()):
        inp = inputs.nth(i)
        val = inp.get_attribute("value") or inp.inner_text()
        inp_id = inp.get_attribute("id") or inp.get_attribute("name")
        if val:
            print(f"Input [{inp_id}] = '{val}'")
            
    ctx.close()
