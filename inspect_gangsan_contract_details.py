import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GANGSAN_DIR = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리\01_임대차계약\22_강남_도곡로1길23_지하1층,1층,2층,3층")

print("Inspecting Gangsan Construction (강산건설) contract PDFs in:", GANGSAN_DIR)

pdf_files = [f for f in os.listdir(GANGSAN_DIR) if f.endswith(".pdf")]
print("Found PDFs:", pdf_files)

for f in pdf_files:
    fp = os.path.join(GANGSAN_DIR, f)
    print(f"\n==========================================")
    print(f"File: {f}")
    print(f"==========================================")
    try:
        doc = fitz.open(fp)
        for page_num, page in enumerate(doc, 1):
            text = page.get_text()
            print(f"--- Page {page_num} ---")
            lines = text.split("\n")
            for l in lines:
                l_str = l.strip()
                if any(k in l_str for k in ["보증금", "금", "원", "삼억", "오억", "십억", "300", "500", "000", "차임", "월세", "임대료"]):
                    print("  ", l_str)
    except Exception as e:
        print("  Error:", e)
