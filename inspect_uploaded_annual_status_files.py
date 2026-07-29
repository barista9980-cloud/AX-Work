import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
UPLOAD_DIR = os.path.join(REAL_ESTATE_BASE, "00_연도별_자산현황_자료")

print("Checking uploaded files in:", UPLOAD_DIR)

if os.path.exists(UPLOAD_DIR):
    uploaded_files = os.listdir(UPLOAD_DIR)
    print(f"Total uploaded files in 00_연도별_자산현황_자료: {len(uploaded_files)}")
    for f in uploaded_files:
        fp = os.path.join(UPLOAD_DIR, f)
        size_kb = os.path.getsize(fp) / 1024
        print(f"  - [FILE] {f} ({size_kb:.1f} KB)")

# Also check desktop / local directory just in case
local_dir = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
print("\nChecking uploaded files in local directory:", local_dir)
for f in os.listdir(local_dir):
    if "22" in f or "23" in f or "24" in f or "25" in f or "현황" in f or "리스트" in f or f.endswith(".pdf"):
        fp = os.path.join(local_dir, f)
        if os.path.isfile(fp):
            size_kb = os.path.getsize(fp) / 1024
            print(f"  - [FILE] {f} ({size_kb:.1f} KB)")
