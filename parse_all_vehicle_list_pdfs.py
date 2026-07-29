import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리")
UPLOAD_DIR = os.path.join(VEHICLE_BASE, "00_차량_업로드_자료")

annual_folders = sorted(os.listdir(UPLOAD_DIR))

for folder in annual_folders:
    folder_p = os.path.join(UPLOAD_DIR, folder)
    if os.path.isdir(folder_p):
        print(f"\n==========================================")
        print(f"=== VEHICLE ANNUAL ASSET FOLDER: {folder} ===")
        print(f"==========================================")
        
        pdf_files = [f for f in os.listdir(folder_p) if f.lower().endswith(".pdf")]
        summary_pdfs = [f for f in pdf_files if f.startswith("0.") or "리스트" in f or "목록" in f]
        
        for s_pdf in summary_pdfs:
            sp_path = os.path.join(folder_p, s_pdf)
            print(f"\n--- Extracting Vehicle Summary PDF: {s_pdf} ---")
            doc = fitz.open(sp_path)
            for p_idx, page in enumerate(doc):
                txt = page.get_text()
                print(f"Page {p_idx+1} Text:\n{txt[:2500]}")
