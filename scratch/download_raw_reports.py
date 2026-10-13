import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission/Details/08VuIoFsBaXGehnOxJ3PMQ__', wait_until='domcontentloaded')
    time.sleep(2)
    
    # Try fetching the compiled raw data report
    # Let's see what endpoint works
    print("Testing SubmissionCompiledRawDataReportPdf...")
    resp = page.request.post('https://etrax.eagleanalytical.com/Report/SubmissionCompiledRawDataReportPdf', 
                             form={'id': '08VuIoFsBaXGehnOxJ3PMQ__'})
    print(f"Status: {resp.status}, Content-Type: {resp.headers.get('content-type')}, Length: {len(resp.body())}")
    if resp.status == 200 and 'pdf' in resp.headers.get('content-type', '').lower():
        with open('scratch/submission_raw_data_report.pdf', 'wb') as f:
            f.write(resp.body())
        print("Successfully saved scratch/submission_raw_data_report.pdf!")
        
    print("\nTesting SubmissionTestCompiledRawDataReportPdf...")
    resp2 = page.request.post('https://etrax.eagleanalytical.com/Report/SubmissionTestCompiledRawDataReportPdf',
                              form={'id': 'zFdCKB28Sd0gPid31SfEnA__'})
    print(f"Status: {resp2.status}, Content-Type: {resp2.headers.get('content-type')}, Length: {len(resp2.body())}")
    if resp2.status == 200 and 'pdf' in resp2.headers.get('content-type', '').lower():
        with open('scratch/submission_test_raw_data_report.pdf', 'wb') as f:
            f.write(resp2.body())
        print("Successfully saved scratch/submission_test_raw_data_report.pdf!")
        
    ctx.close()
