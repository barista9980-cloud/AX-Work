import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
UPLOAD_DIR = os.path.join(REAL_ESTATE_BASE, "00_연도별_자산현황_자료")

print("Inspecting details of items in:", UPLOAD_DIR)

for item in os.listdir(UPLOAD_DIR):
    fp = os.path.join(UPLOAD_DIR, item)
    is_dir = os.path.isdir(fp)
    print(f"\nItem: {item} | Is Dir: {is_dir}")
    if is_dir:
        for root, dirs, files in os.walk(fp):
            print(f"  Subroot: {root}")
            for d in dirs:
                print(f"    Subdir: {d}")
            for f in files:
                f_full = os.path.join(root, f)
                sz = os.path.getsize(f_full) / 1024
                print(f"    File: {f} ({sz:.1f} KB)")
