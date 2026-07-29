import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GENERAL_AFFAIRS_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무")
INSURANCE_UPLOAD_DIR = os.path.join(GENERAL_AFFAIRS_BASE, r"03_보험_자산관리\00_보험_업로드_자료")

print("Checking uploaded files in insurance upload directory:", INSURANCE_UPLOAD_DIR)

if os.path.exists(INSURANCE_UPLOAD_DIR):
    for root, dirs, files in os.walk(INSURANCE_UPLOAD_DIR):
        print(f"\nRoot: {root}")
        for d in dirs:
            print(f"  [DIR] {d}")
        for f in files:
            fp = os.path.join(root, f)
            sz = os.path.getsize(fp) / 1024
            print(f"  [FILE] {f} ({sz:.1f} KB)")

# Also check top-level 03_기업보험_안전관리 and 03_보험_자산관리
for ins_dir_name in ["03_기업보험_안전관리", "03_보험_자산관리"]:
    ins_dir = os.path.join(GENERAL_AFFAIRS_BASE, ins_dir_name)
    print(f"\nChecking all items in: {ins_dir}")
    if os.path.exists(ins_dir):
        for item in os.listdir(ins_dir):
            if item != "00_보험_업로드_자료":
                fp = os.path.join(ins_dir, item)
                print(f"  - [{'DIR' if os.path.isdir(fp) else 'FILE'}] {item}")
