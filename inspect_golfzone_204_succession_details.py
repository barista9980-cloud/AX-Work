import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리\01_임대차계약\02_대전_골프존_204호,상담실")

print("Inspecting Golfzone 204 Succession Contract PDF & Details...")

pdf_files = [f for f in os.listdir(REAL_ESTATE_BASE) if f.endswith(".pdf")]
print("Found PDFs:", pdf_files)

for pdf_name in pdf_files:
    pdf_p = os.path.join(REAL_ESTATE_BASE, pdf_name)
    print(f"\n--- Reading PDF: {pdf_name} ---")
    doc = fitz.open(pdf_p)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    print(full_text[:2000])
