import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_CONTRACTS_DIR = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리\01_차량계약_리스_렌트")

print("Checking existing vehicle contract files in:", VEHICLE_CONTRACTS_DIR)

if os.path.exists(VEHICLE_CONTRACTS_DIR):
    for root, dirs, files in os.walk(VEHICLE_CONTRACTS_DIR):
        print(f"\nSubroot: {root}")
        for f in files:
            fp = os.path.join(root, f)
            sz_kb = os.path.getsize(fp) / 1024
            print(f"  - File: {f} ({sz_kb:.1f} KB)")
