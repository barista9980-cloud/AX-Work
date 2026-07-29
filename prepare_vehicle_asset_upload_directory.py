import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리")
UPLOAD_DIR = os.path.join(VEHICLE_BASE, "00_차량_업로드_자료")

os.makedirs(UPLOAD_DIR, exist_ok=True)
print("Vehicle upload directory prepared:", UPLOAD_DIR)

print("\nExisting items in 02_차량_자산관리:")
if os.path.exists(VEHICLE_BASE):
    for item in os.listdir(VEHICLE_BASE):
        fp = os.path.join(VEHICLE_BASE, item)
        is_dir = os.path.isdir(fp)
        print(f"  - [{'DIR' if is_dir else 'FILE'}] {item}")
