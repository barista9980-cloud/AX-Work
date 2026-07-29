import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
UPLOAD_DIR = os.path.join(REAL_ESTATE_BASE, "00_연도별_자산현황_자료")

os.makedirs(UPLOAD_DIR, exist_ok=True)
print("Upload directory prepared:", UPLOAD_DIR)
