import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_1TO1_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리\01_차량계약_리스_렌트")

print("Inspecting Signed Contract PDF Files for EXACT Deposit & Advance Amounts...")

v_dirs = sorted([d for d in os.listdir(VEHICLE_1TO1_BASE) if os.path.isdir(os.path.join(VEHICLE_1TO1_BASE, d))])

for v_dir in v_dirs:
    dir_p = os.path.join(VEHICLE_1TO1_BASE, v_dir)
    pdf_files = [f for f in os.listdir(dir_p) if f.endswith(".pdf")]
    
    print(f"\n==========================================")
    print(f"Folder: {v_dir}")
    print(f"==========================================")
    
    for pdf_name in pdf_files:
        pdf_p = os.path.join(dir_p, pdf_name)
        print("  PDF:", pdf_name)
        try:
            doc = fitz.open(pdf_p)
            for page_num, page in enumerate(doc, 1):
                text = page.get_text()
                lines = [l.strip() for l in text.split("\n") if l.strip()]
                for l in lines:
                    if any(k in l for k in ["보증금", "선납금", "보증율", "납입보증금", "보증금액", "원금", "취득원가"]):
                        print(f"    [P{page_num}] {l}")
        except Exception as e:
            print("    Error reading PDF:", e)
