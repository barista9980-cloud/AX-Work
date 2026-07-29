import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리")
UPLOAD_DIR = os.path.join(VEHICLE_BASE, "00_차량_업로드_자료")

print("Checking uploaded files in vehicle upload directory:", UPLOAD_DIR)

if os.path.exists(UPLOAD_DIR):
    for root, dirs, files in os.walk(UPLOAD_DIR):
        print(f"\nRoot: {root}")
        for d in dirs:
            print(f"  [DIR] {d}")
        for f in files:
            fp = os.path.join(root, f)
            sz = os.path.getsize(fp) / 1024
            print(f"  [FILE] {f} ({sz:.1f} KB)")
else:
    print("Upload directory not found:", UPLOAD_DIR)

# Also check root VEHICLE_BASE for any new folders
print("\nChecking top-level folders in 02_차량_자산관리:")
for item in os.listdir(VEHICLE_BASE):
    fp = os.path.join(VEHICLE_BASE, item)
    if os.path.isdir(fp):
        print(f"  [TOP DIR] {item}")
