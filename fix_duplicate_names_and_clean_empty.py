import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
LEASE_DIR = os.path.join(REAL_ESTATE_BASE, "01_임대차계약")

print("Fixing duplicate folder names and cleaning leftover empty folders...")

# 1. Remove leftover old empty folder 01_가산_대륭포스트타워6차
old_gasan = os.path.join(LEASE_DIR, "01_가산_대륭포스트타워6차")
if os.path.exists(old_gasan):
    try:
        shutil.rmtree(old_gasan)
        print("  [DELETED LEFTOVER EMPTY FOLDER] 01_가산_대륭포스트타워6차")
    except Exception as e:
        print("  [ERROR DELETING] 01_가산_대륭포스트타워6차:", e)

# 2. Fix 01_판교_판교동612_판교동612 -> 01_판교_판교동612
pangyo_double = os.path.join(LEASE_DIR, "01_판교_판교동612_판교동612")
pangyo_clean = os.path.join(LEASE_DIR, "01_판교_판교동612")

if os.path.exists(pangyo_double):
    try:
        os.rename(pangyo_double, pangyo_clean)
        print("  [FIXED FOLDER NAME] 01_판교_판교동612_판교동612 -> 01_판교_판교동612")
        
        # Rename docx note inside if needed
        old_note = os.path.join(pangyo_clean, "부동산_계약관리노트_판교_판교동612_판교동612.docx")
        new_note = os.path.join(pangyo_clean, "부동산_계약관리노트_판교_판교동612.docx")
        if os.path.exists(old_note):
            os.rename(old_note, new_note)
            print("    - [FIXED NOTE NAME] 부동산_계약관리노트_판교_판교동612.docx")
    except Exception as e:
        print("  [ERROR FIXING PANGYO]:", e)

# 3. Fix 10_대전_도룡동385-28_385-28 -> 10_대전_도룡동385-28
doryong_double = os.path.join(LEASE_DIR, "10_대전_도룡동385-28_385-28")
doryong_clean = os.path.join(LEASE_DIR, "10_대전_도룡동385-28")

if os.path.exists(doryong_double):
    try:
        os.rename(doryong_double, doryong_clean)
        print("  [FIXED FOLDER NAME] 10_대전_도룡동385-28_385-28 -> 10_대전_도룡동385-28")

        # Rename docx note inside if needed
        old_note = os.path.join(doryong_clean, "부동산_계약관리노트_대전_도룡동385-28_385-28.docx")
        new_note = os.path.join(doryong_clean, "부동산_계약관리노트_대전_도룡동385-28.docx")
        if os.path.exists(old_note):
            os.rename(old_note, new_note)
            print("    - [FIXED NOTE NAME] 부동산_계약관리노트_대전_도룡동385-28.docx")
    except Exception as e:
        print("  [ERROR FIXING DORYONG]:", e)

print("\n=== FINAL VERIFIED LIST OF FOLDERS IN 01_임대차계약 ===")
final_folders = sorted(os.listdir(LEASE_DIR))
for idx, f in enumerate(final_folders, 1):
    print(f"  {idx:02d}. {f}")
