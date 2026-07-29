import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

trash_gasan = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\_DEL_01_가산_대륭포스트타워6차"

if os.path.exists(trash_gasan):
    try:
        shutil.rmtree(trash_gasan)
        print("Successfully deleted trash gasan folder!")
    except Exception as e:
        print("Error deleting trash gasan:", e)

LEASE_DIR = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약"
print("\nFinal Clean Folders in 01_임대차계약:")
for item in sorted(os.listdir(LEASE_DIR)):
    if not item.startswith("_"):
        print(" -", item)
