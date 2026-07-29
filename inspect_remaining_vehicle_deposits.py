import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_1TO1_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리\01_차량계약_리스_렌트")

v_dirs = sorted([d for d in os.listdir(VEHICLE_1TO1_BASE) if os.path.isdir(os.path.join(VEHICLE_1TO1_BASE, d))])

for idx, v_dir in enumerate(v_dirs, 1):
    dir_p = os.path.join(VEHICLE_1TO1_BASE, v_dir)
    pdf_files = [f for f in os.listdir(dir_p) if f.endswith(".pdf")]
    if pdf_files:
        pdf_p = os.path.join(dir_p, pdf_files[0])
        doc = fitz.open(pdf_p)
        full_text = ""
        for page in doc:
            full_text += page.get_text() + "\n"
        print(f"\n==========================================")
        print(f"[{idx:02d}] {v_dir}")
        print(f"==========================================")
        # Search lines with 보증금, 선납금, 리스료, 렌트료
        for line in full_text.split("\n"):
            line_str = line.strip()
            if any(k in line_str for k in ["보증금", "선납금", "선수금", "보증율", "취득원가", "잔존가치", "매회납부"]):
                print("  ", line_str)
