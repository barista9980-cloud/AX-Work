import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

target_base = None
for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    if "01_부동산_자산관리" in root:
        target_base = root
        break

print("Found 01_부동산_자산관리 at:", target_base)

if target_base:
    print("\n--- CURRENT SUBFOLDERS & FILES ---")
    for root, dirs, files in os.walk(target_base):
        rel_p = os.path.relpath(root, target_base)
        print(f"\n[DIR] {rel_p}")
        for f in files:
            print(f"   - {f}")
