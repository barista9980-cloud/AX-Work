import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GENERAL_AFFAIRS_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무")

DIR_SAFETY = os.path.join(GENERAL_AFFAIRS_BASE, "03_기업보험_안전관리")
DIR_INSURANCE = os.path.join(GENERAL_AFFAIRS_BASE, "03_보험_자산관리")

print("Consolidating insurance folders to 03_보험_자산관리...")
print("Source 1 (DIR_SAFETY):", DIR_SAFETY)
print("Target (DIR_INSURANCE):", DIR_INSURANCE)

os.makedirs(DIR_INSURANCE, exist_ok=True)

# STEP 1: Move all 1:1 insurance flat folders from DIR_SAFETY to DIR_INSURANCE
if os.path.exists(DIR_SAFETY):
    for item in os.listdir(DIR_SAFETY):
        src_path = os.path.join(DIR_SAFETY, item)
        dst_path = os.path.join(DIR_INSURANCE, item)
        
        if os.path.isdir(src_path):
            if os.path.exists(dst_path):
                shutil.rmtree(dst_path, ignore_errors=True)
            shutil.move(src_path, dst_path)
            print(f"  [MOVED DIR] {item} -> 03_보험_자산관리/{item}")
        elif os.path.isfile(src_path):
            shutil.copy2(src_path, dst_path)
            os.remove(src_path)
            print(f"  [MOVED FILE] {item} -> 03_보험_자산관리/{item}")

    # Remove DIR_SAFETY completely
    try:
        shutil.rmtree(DIR_SAFETY, ignore_errors=True)
        print("  [REMOVED OLD FOLDER] 03_기업보험_안전관리 completely removed!")
    except Exception as e:
        print("  [ERROR REMOVING DIR_SAFETY]", e)

# STEP 2: Update all script targets so DIR_INSURANCE is used
print("\n--- Listing final consolidated structure in 03_보험_자산관리 ---")
for item in sorted(os.listdir(DIR_INSURANCE)):
    fp = os.path.join(DIR_INSURANCE, item)
    is_dir = os.path.isdir(fp)
    print(f"  - [{'DIR' if is_dir else 'FILE'}] {item}")

print("\n==========================================")
print("INSURANCE DIRECTORY CONSOLIDATION TO 03_보험_자산관리 COMPLETE!")
print("==========================================")
