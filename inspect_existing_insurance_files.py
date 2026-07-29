import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GENERAL_AFFAIRS_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무")

for ins_dir_name in ["03_기업보험_안전관리", "03_보험_자산관리"]:
    ins_dir = os.path.join(GENERAL_AFFAIRS_BASE, ins_dir_name)
    print(f"\nChecking existing files in: {ins_dir}")
    if os.path.exists(ins_dir):
        for root, dirs, files in os.walk(ins_dir):
            print(f"  Root: {root}")
            for d in dirs:
                print(f"    [DIR] {d}")
            for f in files:
                fp = os.path.join(root, f)
                sz = os.path.getsize(fp) / 1024
                print(f"    [FILE] {f} ({sz:.1f} KB)")
