import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")

golfzone_files = []

for root, dirs, files in os.walk(REAL_ESTATE_BASE):
    if "골프존" in root:
        for f in files:
            if f.endswith(".pdf"):
                golfzone_files.append((root, f))

golfzone_files.sort(key=lambda x: x[1])

print("=== GOLFZONE PDF CONTRACT FILES ===")
for r, f in golfzone_files:
    folder_name = os.path.basename(r)
    print(f"Folder: [{folder_name}] | File: {f}")
