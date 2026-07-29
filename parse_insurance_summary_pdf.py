import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
INSURANCE_UPLOAD_DIR = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\03_보험_자산관리\00_보험_업로드_자료")

summary_pdf_path = os.path.join(INSURANCE_UPLOAD_DIR, r"2025_폭스에듀_보험계약서\0.폭스에듀_보험증권_계약리스트_251124.pdf")

print("Extracting Insurance Summary List PDF:", summary_pdf_path)

if os.path.exists(summary_pdf_path):
    doc = fitz.open(summary_pdf_path)
    for p_idx, page in enumerate(doc):
        txt = page.get_text()
        print(f"\n--- Page {p_idx+1} Text ---\n{txt}")
else:
    print("Summary PDF not found!")
