import os
from playwright.sync_api import sync_playwright

user_data_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
url_test = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/ipbSVhVxCkfjT1rdWrMzOg__#Worksheets'

with sync_playwright() as p:
    browser = p.chromium.launch_persistent_context(user_data_dir=user_data_dir, headless=True)
    page = browser.new_page()
    page.goto(url_test, wait_until='domcontentloaded')
    page.wait_for_timeout(3000)
    
    # Check if we need to enter username
    if "Log in" in page.title():
        page.fill("input[name='Username']", "qchen")
        page.click("input[type='submit'], button[type='submit']")
        page.wait_for_timeout(3000)
        
    print("Page URL:", page.url)
    
    # Let's inspect the Worksheets and Test Results section
    # Let's take a screenshot of the Worksheets tab and Test Details tab
    page.screenshot(path="scratch/page_view.png", full_page=True)
    
    # Evaluate JS to extract all fieldsets, tables, form controls
    res = page.evaluate('''() => {
        const results = [];
        const inputs = document.querySelectorAll('input, select, textarea');
        inputs.forEach(el => {
            let label = '';
            // find associated label
            if (el.id) {
                const l = document.querySelector('label[for="' + el.id + '"]');
                if (l) label = l.innerText.trim();
            }
            if (!label) {
                const parent = el.closest('.form-group') || el.closest('tr') || el.parentElement;
                if (parent) {
                    const l = parent.querySelector('label, th, td:first-child');
                    if (l) label = l.innerText.trim();
                }
            }
            const val = el.tagName === 'SELECT' ? (el.options[el.selectedIndex]?.text || el.value) : el.value;
            results.push({
                id: el.id,
                name: el.name,
                label: label,
                value: val,
                type: el.type,
                tag: el.tagName
            });
        });
        return results;
    }''')
    
    for r in res:
        if r['value']:
            print(f"[{r['tag']}] id={r['id']} | name={r['name']} | label={r['label']} => VALUE: {r['value']}")
            
    # Also look for any spans/divs with test result values
    text_blocks = page.evaluate('''() => {
        const out = [];
        document.querySelectorAll('#TestResults dl, #TestResults table, #TestResults .form-horizontal, #worksheets, .tab-pane').forEach(el => {
            out.push({id: el.id, class: el.className, text: el.innerText});
        });
        return out;
    }''')
    print("\n--- TAB PANES / TEST RESULTS TEXT ---")
    for tb in text_blocks:
        print(f"ID: {tb['id']} | CLASS: {tb['class']}")
        print(tb['text'][:500])
        print("="*40)
        
    browser.close()
