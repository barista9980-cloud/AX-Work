import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GENERAL_AFFAIRS_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무")

upload_queue_p = os.path.join(GENERAL_AFFAIRS_BASE, "06_자동파싱_업로드큐")

print("Checking [총무]업무 Upload Queue Directory...")

os.makedirs(upload_queue_p, exist_ok=True)

# Sub-queues for guidance (optional)
sub_queues = ["01_부동산_대기", "02_차량_대기", "03_보험_대기", "04_일반_대기"]
for sq in sub_queues:
    os.makedirs(os.path.join(upload_queue_p, sq), exist_ok=True)

print("  Upload Queue Directory Ready:", upload_queue_p)
for item in os.listdir(upload_queue_p):
    print("   -", item)
