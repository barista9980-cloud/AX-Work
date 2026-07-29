import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
INSURANCE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\03_보험_자산관리")

print("Cleaning unused subdirectories in 03_보험_자산관리...")

for item in os.listdir(INSURANCE_BASE):
    fp = os.path.join(INSURANCE_BASE, item)
    if os.path.isdir(fp):
        # If it's one of the old category folders with no files, remove it
        if item in ["01_화재_배상책임보험", "02_임직원_단체보험", "03_중대재해_안전보건"]:
            shutil.rmtree(fp, ignore_errors=True)
            print(f"  [REMOVED UNUSED FOLDER] {item}")

print("\n--- Final Clean Structure of 03_보험_자산관리 ---")
for item in sorted(os.listdir(INSURANCE_BASE)):
    fp = os.path.join(INSURANCE_BASE, item)
    is_dir = os.path.isdir(fp)
    print(f"  - [{'DIR' if is_dir else 'FILE'}] {item}")
