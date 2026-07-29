import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
INSURANCE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\03_보험_자산관리")

print("Executing Refined Integrated Insurance Structure in 03_보험_자산관리...")

# Categories to create
cat_folders = {
    "00_연도별_가입현황_비교견적": os.path.join(INSURANCE_BASE, "00_연도별_가입현황_비교견적"),
    "01_보험증권_및_배서계약서": os.path.join(INSURANCE_BASE, "01_보험증권_및_배서계약서"),
    "02_보험금청구_사고접수": os.path.join(INSURANCE_BASE, "02_보험금청구_사고접수"),
    "03_보험료납입_증빙": os.path.join(INSURANCE_BASE, "03_보험료납입_증빙")
}

for k, p in cat_folders.items():
    os.makedirs(p, exist_ok=True)

# Move upload folder inside 00_연도별_가입현황_비교견적
upload_src = os.path.join(INSURANCE_BASE, "00_보험_업로드_자료")
upload_dst = os.path.join(cat_folders["00_연도별_가입현황_비교견적"], "00_보험_업로드_자료")
if os.path.exists(upload_src) and not os.path.exists(upload_dst):
    shutil.move(upload_src, upload_dst)
    print("  [MOVED UPLOAD DIR] 00_보험_업로드_자료 -> 00_연도별_가입현황_비교견적/00_보험_업로드_자료")

# Move all 1:1 policy folders (01_ ~ 12_) into 01_보험증권_및_배서계약서
dir_01_target = cat_folders["01_보험증권_및_배서계약서"]

for item in os.listdir(INSURANCE_BASE):
    fp = os.path.join(INSURANCE_BASE, item)
    if os.path.isdir(fp) and item not in cat_folders and item != "00_보험_업로드_자료":
        dst_p = os.path.join(dir_01_target, item)
        if os.path.exists(dst_p):
            shutil.rmtree(dst_p, ignore_errors=True)
        shutil.move(fp, dst_p)
        print(f"  [INTEGRATED 1:1 POLICY FOLDER] {item} -> 01_보험증권_및_배서계약서/{item}")

print("\n--- Final Refined Structure in 03_보험_자산관리 ---")
for main_item in sorted(os.listdir(INSURANCE_BASE)):
    main_fp = os.path.join(INSURANCE_BASE, main_item)
    if os.path.isdir(main_fp):
        print(f"\n📂 [{main_item}]")
        for sub_item in sorted(os.listdir(main_fp)):
            sub_fp = os.path.join(main_fp, sub_item)
            is_sub_dir = os.path.isdir(sub_fp)
            print(f"   - [{'DIR' if is_sub_dir else 'FILE'}] {sub_item}")

print("\n==========================================")
print("REFINED INTEGRATED INSURANCE STRUCTURE COMPLETE!")
print("==========================================")
