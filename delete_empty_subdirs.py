import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

old_gasan = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_가산_대륭포스트타워6차"

# Try removing subdirectories first
sub1510 = os.path.join(old_gasan, "1510호")
if os.path.exists(sub1510):
    try:
        os.rmdir(sub1510)
        print("Successfully removed empty subfolder 1510호!")
    except Exception as e:
        print("Error removing 1510호:", e)

if os.path.exists(old_gasan):
    try:
        os.rmdir(old_gasan)
        print("Successfully removed old parent folder 01_가산_대륭포스트타워6차!")
    except Exception as e:
        print("Error removing old parent folder:", e)
