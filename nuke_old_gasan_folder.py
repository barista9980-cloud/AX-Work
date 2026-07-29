import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

old_gasan = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_가산_대륭포스트타워6차"

if os.path.exists(old_gasan):
    print("Deleting old empty folder:", old_gasan)
    cmd = f'rmdir /s /q "{old_gasan}"'
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if not os.path.exists(old_gasan):
        print("  [SUCCESSFULLY REMOVED OLD 01_가산_대륭포스트타워6차 FOLDER]")
    else:
        print("  [STILL EXISTS, RETRYING...] STDERR:", res.stderr)
else:
    print("Folder does not exist:", old_gasan)
