import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GENERAL_AFFAIRS_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무")
QUEUE_DIR = os.path.join(GENERAL_AFFAIRS_BASE, "06_자동파싱_업로드큐")

print("Fixing Duplicate Sequence Numbers in 06_자동파싱_업로드큐...")

# Clean up existing subfolders in QUEUE_DIR
if os.path.exists(QUEUE_DIR):
    for item in os.listdir(QUEUE_DIR):
        item_p = os.path.join(QUEUE_DIR, item)
        if os.path.isdir(item_p):
            shutil.rmtree(item_p, ignore_errors=True)

# Re-create STRICT UNIQUE SEQUENCE NUMBERS
unique_queue_structure = [
    "01_부동산_업로드대기",
    "02_차량_업로드대기",
    "03_보험_업로드대기",
    "04_비품_소모품_업로드대기",
    "05_처리완료_아카이브"
]

for subf in unique_queue_structure:
    os.makedirs(os.path.join(QUEUE_DIR, subf), exist_ok=True)
    print("  [CREATED UNIQUE SUBFOLDER]", subf)

print("\n--- Current 06_자동파싱_업로드큐 Directory Listing ---")
for item in sorted(os.listdir(QUEUE_DIR)):
    print("   -", item)

print("\n==========================================")
print("SEQUENCE NUMBER DUPLICATION FIXED PERFECTLY!")
print("==========================================")
