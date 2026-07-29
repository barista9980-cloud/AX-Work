import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")

print("Checking ALL folders under:", REAL_ESTATE_BASE)

for root, dirs, files in os.walk(REAL_ESTATE_BASE):
    # Only list direct children of top folders
    rel = os.path.relpath(root, REAL_ESTATE_BASE)
    print(f"\nDirectory: [{rel}]")
    for d in dirs:
        print(f"  - [DIR] {d}")
