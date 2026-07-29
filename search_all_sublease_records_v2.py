import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
UPLOAD_DIR = os.path.join(REAL_ESTATE_BASE, "00_연도별_자산현황_자료")

print("Searching sublease (전대차) files in:", UPLOAD_DIR)

sublease_files = []

for root, dirs, files in os.walk(UPLOAD_DIR):
    for f in files:
        if f.lower().endswith(".pdf"):
            if "전대" in f or "전대차" in f or "전차인" in f:
                fp = os.path.join(root, f)
                sublease_files.append((f, fp, root))

print(f"\n==========================================")
print(f"Total Sublease (전대차) Files Found in Uploaded Directories: {len(sublease_files)}")
print(f"==========================================")

for idx, (fname, fpath, root_dir) in enumerate(sublease_files, 1):
    annual_folder = os.path.basename(root_dir)
    print(f"\n[{idx}] Folder: {annual_folder} | File: {fname}")
    try:
        doc = fitz.open(fpath)
        txt = doc[0].get_text() if len(doc) > 0 else ""
        lines = [l.strip() for l in txt.split("\n") if l.strip()]
        print("  Snippet:", " | ".join(lines[:10]))
    except Exception as e:
        print("  Error reading PDF:", e)
