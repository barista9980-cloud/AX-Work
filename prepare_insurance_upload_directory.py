import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GENERAL_AFFAIRS_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무")
INSURANCE_BASE = os.path.join(GENERAL_AFFAIRS_BASE, "03_보험_자산관리")
UPLOAD_DIR = os.path.join(INSURANCE_BASE, "00_보험_업로드_자료")

os.makedirs(UPLOAD_DIR, exist_ok=True)
print("Insurance upload directory prepared:", UPLOAD_DIR)

print("\nExisting directories in [총무]업무:")
if os.path.exists(GENERAL_AFFAIRS_BASE):
    for item in os.listdir(GENERAL_AFFAIRS_BASE):
        fp = os.path.join(GENERAL_AFFAIRS_BASE, item)
        is_dir = os.path.isdir(fp)
        print(f"  - [{'DIR' if is_dir else 'FILE'}] {item}")
