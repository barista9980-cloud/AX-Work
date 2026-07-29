import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

dogok_path = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\02_강남_도곡로1길23\지하1층,1층,2층,3층"

print("Files currently in Dogok folder:")
for f in os.listdir(dogok_path):
    print(" -", f)
