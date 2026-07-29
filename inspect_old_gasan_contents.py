import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

old_gasan = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_가산_대륭포스트타워6차"

print("Contents of old_gasan:")
for root, dirs, files in os.walk(old_gasan):
    print("Root:", root)
    for d in dirs:
        print("  Subdir:", d)
    for f in files:
        print("  File:", f)
