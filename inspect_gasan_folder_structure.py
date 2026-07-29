import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
LEASE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리\01_임대차계약")

gasan_path = None
for item in os.listdir(LEASE_BASE):
    if "대륭포스트타워6차" in item or "가산" in item:
        gasan_path = os.path.join(LEASE_BASE, item)
        break

print("Gasan folder path found:", gasan_path)

if gasan_path:
    print("\nContents of Gasan folder:")
    for item in os.listdir(gasan_path):
        full_sub = os.path.join(gasan_path, item)
        if os.path.isdir(full_sub):
            print(f" [SUBDIR] {item}")
            for f in os.listdir(full_sub):
                print(f"     - {f}")
        else:
            print(f" [FILE] {item}")
