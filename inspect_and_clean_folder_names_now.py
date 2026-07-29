import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
LEASE_DIR = os.path.join(REAL_ESTATE_BASE, "01_임대차계약")
SALE_DIR = os.path.join(REAL_ESTATE_BASE, "02_매매_소유권문서")

print("=== INSPECTING ALL FOLDERS IN 01_임대차계약 ===")
lease_folders = sorted(os.listdir(LEASE_DIR))
for f in lease_folders:
    fp = os.path.join(LEASE_DIR, f)
    if os.path.isdir(fp):
        print(f" [LEASE DIR] {f}")

print("\n=== INSPECTING ALL FOLDERS IN 02_매매_소유권문서 ===")
sale_folders = sorted(os.listdir(SALE_DIR))
for f in sale_folders:
    fp = os.path.join(SALE_DIR, f)
    if os.path.isdir(fp):
        print(f" [SALE DIR] {f}")
