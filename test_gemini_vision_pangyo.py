import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\01_최초임대차_판교_판교동612_[박동석-㈜폭스에듀]_(210731).pdf"

doc = fitz.open(pdf_path)
full_text = ""
for page in doc:
    full_text += page.get_text() + "\n"

print("--- FULL EXTRACTED TEXT FROM PANGYO PDF ---")
print(full_text)
