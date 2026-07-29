import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")
UPLOAD_DIR = os.path.join(REAL_ESTATE_BASE, "00_연도별_자산현황_자료")

if not os.path.exists(UPLOAD_DIR):
    UPLOAD_DIR = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리\00_연도별_부동산_총괄자산대장\00_연도별_자산현황_자료")

print("Deep searching ALL uploaded PDFs for sublease (전대차) records...")

sublease_findings = []

for root, dirs, files in os.walk(UPLOAD_DIR):
    for f in files:
        if f.lower().endswith(".pdf"):
            fp = os.path.join(root, f)
            try:
                doc = fitz.open(fp)
                full_txt = ""
                for page in doc:
                    full_txt += page.get_text() + "\n"
                
                if "전대" in full_txt or "전차인" in f or "전대차" in f or "전대" in f:
                    sublease_findings.append({
                        "filename": f,
                        "path": fp,
                        "text_snippet": full_txt[:1000]
                    })
            except Exception:
                pass

print(f"\n==========================================")
print(f"TOTAL SUBLEASE (전대차) FILES FOUND: {len(sublease_findings)}")
print(f"==========================================")

for idx, item in enumerate(sublease_findings, 1):
    print(f"\n[{idx}] File: {item['filename']}")
    # print snippet lines containing 전대
    lines = item['text_snippet'].split("\n")
    for line in lines:
        if "전대" in line or "전차" in line or "계약" in line or "보증금" in line or "월" in line:
            print(f"    {line.strip()}")
