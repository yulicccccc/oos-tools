import os, sys
from playwright.sync_api import sync_playwright

user_data_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
url_test = "https://etrax.eagleanalytical.com/SubmissionTest/Details/ipbSVhVxCkfjT1rdWrMzOg__#Worksheets"

with sync_playwright() as p:
    browser = p.chromium.launch_persistent_context(user_data_dir=user_data_dir, headless=True)
    page = browser.new_page()
    page.goto(url_test, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    print("Initial Page title:", page.title())
    print("Initial Page url:", page.url)
    
    # Check if login form is present
    if "Log in" in page.title() or "Sign in" in page.title() or "login" in page.url.lower():
        print("Handling login...")
        try:
            # Step 1: Username input
            user_input = page.locator("input[name='Username'], input#Username, input[type='text']")
            if user_input.is_visible():
                print("Entering username...")
                user_input.fill("qchen")
                # Look for Continue button
                submit_btn = page.locator("input[type='submit'], button[type='submit'], button:has-text('Continue')")
                submit_btn.click()
                page.wait_for_timeout(4000)
                print("After continue, title:", page.title(), "url:", page.url)
                
            # Check if redirected to Microsoft SSO or password field
            if "login.microsoftonline.com" in page.url or "Sign in" in page.title():
                print("At Microsoft SSO...")
                if page.locator("input[type='email']").is_visible():
                    page.fill("input[type='email']", "qchen@eagleanalytical.com")
                    page.click("#idSIButton9")
                    page.wait_for_timeout(2000)
                if page.locator("input[type='password']").is_visible():
                    page.fill("input[type='password']", "Xy19910817")
                    page.click("#idSIButton9")
                    page.wait_for_timeout(3000)
                # Check Stay signed in?
                if page.locator("#KmsiCheckboxField").is_visible():
                    page.check("#KmsiCheckboxField")
                    page.click("#idSIButton9")
                    page.wait_for_timeout(3000)
            elif page.locator("input[name='Password'], input[type='password']").is_visible():
                print("Entering password...")
                page.fill("input[name='Password'], input[type='password']", "Xy19910817")
                page.click("input[type='submit'], button[type='submit'], button:has-text('Log in')")
                page.wait_for_timeout(4000)
                print("After login, title:", page.title(), "url:", page.url)
        except Exception as e:
            print("Login error:", e)
            
    # Check current state
    print("Final title:", page.title())
    print("Final url:", page.url)
    page.screenshot(path="scratch/login_step_result.png")
    
    if "SubmissionTest" in page.url:
        print("SUCCESS! Reached SubmissionTest Details page.")
        # Grab all input and select values
        text = page.inner_text("body")
        with open("scratch/submission_test_ipb.txt", "w", encoding="utf-8") as f:
            f.write(text)
        page.screenshot(path="scratch/submission_test_ipb.png")
    else:
        print("Still not on SubmissionTest page.")
        
    browser.close()
