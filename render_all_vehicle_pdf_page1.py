import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_1TO1_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리\01_차량계약_리스_렌트")
output_dir = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\vehicle_page1_imgs"
os.makedirs(output_dir, exist_ok=True)

v_dirs = sorted([d for d in os.listdir(VEHICLE_1TO1_BASE) if os.path.isdir(os.path.join(VEHICLE_1TO1_BASE, d))])

for idx, v_dir in enumerate(v_dirs, 1):
    dir_p = os.path.join(VEHICLE_1TO1_BASE, v_dir)
    pdf_files = [f for f in os.listdir(dir_p) if f.endswith(".pdf")]
    if pdf_files:
        pdf_p = os.path.join(dir_p, pdf_files[0])
        doc = fitz.open(pdf_p)
        page = doc[0]
        pix = page.get_pixmap(dpi=200)
        out_img = os.path.join(output_dir, f"v_{idx:02d}_{pdf_files[0][:30]}.png")
        pix.save(out_img)
        print(f"[{idx:02d}] Saved: {out_img}")

print("All vehicle page 1 images rendered!")
