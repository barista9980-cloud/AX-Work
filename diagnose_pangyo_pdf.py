import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\01_최초임대차_판교_판교동612_[박동석-㈜폭스에듀]_(210731).pdf"

print("Diagnosing PDF file:", pdf_path)
print("File exists:", os.path.exists(pdf_path))

if os.path.exists(pdf_path):
    doc = fitz.open(pdf_path)
    print("Total pages:", len(doc))
    for idx, page in enumerate(doc):
        text = page.get_text()
        print(f"\n--- PAGE {idx+1} (Text Length: {len(text)}) ---")
        print(text[:1000])
