import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

old_gasan = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_가산_대륭포스트타워6차"
trash_gasan = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\_DEL_01_가산_대륭포스트타워6차"

if os.path.exists(old_gasan):
    try:
        os.rename(old_gasan, trash_gasan)
        print("Renamed old gasan to trash_gasan:", trash_gasan)
    except Exception as e:
        print("Error renaming old gasan:", e)

# Print current clean folder list
LEASE_DIR = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약"
print("\nCurrent Folders in 01_임대차계약:")
for item in sorted(os.listdir(LEASE_DIR)):
    print(" -", item)
