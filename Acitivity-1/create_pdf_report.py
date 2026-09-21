"""
Script to generate the executive-grade PDF report for Data Detective Mission (IDA Activity 1).
Takes Data_Detective_EDA_Report.html and prints it to Data_Detective_EDA_Report.pdf using Headless Chrome or Edge.
"""

import os
import subprocess

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_HTML = os.path.join(CURRENT_DIR, "Data_Detective_EDA_Report.html")
OUTPUT_PDF = os.path.join(CURRENT_DIR, "Data_Detective_EDA_Report.pdf")

def render_pdf():
    if not os.path.exists(OUTPUT_HTML):
        print(f"Error: {OUTPUT_HTML} not found.")
        return False
        
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    ]
    chrome_exec = next((p for p in chrome_paths if os.path.exists(p)), None)
    
    if not chrome_exec:
        print("ERROR: No Chrome or Edge executable found to render PDF.")
        return False
        
    print(f"Using browser executable: {chrome_exec}")
    cmd = [
        chrome_exec,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={OUTPUT_PDF}",
        OUTPUT_HTML
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0 and os.path.exists(OUTPUT_PDF):
        size_kb = os.path.getsize(OUTPUT_PDF) / 1024
        print(f"[SUCCESS] PDF successfully generated: {OUTPUT_PDF} ({size_kb:.1f} KB)")
        return True
    else:
        print(f"PDF generation failed with return code {result.returncode}")
        print("Stderr:", result.stderr)
        return False

if __name__ == "__main__":
    render_pdf()
