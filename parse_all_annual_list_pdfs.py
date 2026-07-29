import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
UPLOAD_DIR = os.path.join(REAL_ESTATE_BASE, "00_연도별_자산현황_자료")

annual_folders = sorted(os.listdir(UPLOAD_DIR))

for folder in annual_folders:
    folder_p = os.path.join(UPLOAD_DIR, folder)
    if os.path.isdir(folder_p):
        print(f"\n==========================================")
        print(f"=== ANNUAL ASSET FOLDER: {folder} ===")
        print(f"==========================================")
        
        pdf_files = [f for f in os.listdir(folder_p) if f.lower().endswith(".pdf")]
        
        # Look for summary list PDF first (starts with 0. or contains 리스트)
        summary_pdfs = [f for f in pdf_files if f.startswith("0.") or "리스트" in f or "목록" in f]
        
        print(f"Found {len(summary_pdfs)} Summary List PDFs and {len(pdf_files)} Total Contract PDFs.")
        
        for s_pdf in summary_pdfs:
            sp_path = os.path.join(folder_p, s_pdf)
            print(f"\n--- Extracting Summary PDF: {s_pdf} ---")
            doc = fitz.open(sp_path)
            for p_idx, page in enumerate(doc):
                txt = page.get_text()
                print(f"Page {p_idx+1} Text Snippet:\n{txt[:1500]}")
