import os
import sys
import re
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")

date_proposals = []

for root, dirs, files in os.walk(REAL_ESTATE_BASE):
    pdf_files = [f for f in files if f.lower().endswith(".pdf")]
    if not pdf_files:
        continue

    folder_name = os.path.basename(root)
    parent_name = os.path.basename(os.path.dirname(root))

    pdf_files.sort()

    for f in pdf_files:
        name_no_ext = os.path.splitext(f)[0]
        pdf_path = os.path.join(root, f)

        # Attempt to parse existing 6-digit date
        m_date = re.search(r"\((\d{6})\)", name_no_ext)
        current_date_str = m_date.group(1) if m_date else "미상"

        # Try extracting text from PDF to find actual lease start date
        lease_start_yymmdd = None
        try:
            doc_pdf = fitz.open(pdf_path)
            for page in doc_pdf:
                txt = page.get_text()
                # Look for patterns like 2024년 11월 01일 or 2024.11.01
                m_start = re.search(r"(20\d{2})년?\s*(\d{1,2})월?\s*(\d{1,2})일?\s*부터", txt)
                if not m_start:
                    m_start = re.search(r"(20\d{2})[\.\-년]\s*(\d{1,2})[\.\-월]\s*(\d{1,2})[\.\-일]?", txt)
                if m_start:
                    y, mon, d = m_start.groups()
                    y_2d = y[2:]
                    mon_2d = f"{int(mon):02d}"
                    d_2d = f"{int(d):02d}"
                    lease_start_yymmdd = f"{y_2d}{mon_2d}{d_2d}"
                    break
        except Exception:
            pass

        final_date_str = lease_start_yymmdd if lease_start_yymmdd else current_date_str

        # Reconstruct Option A filename with Lease Start Date
        pattern_a = r"^(\d{2})_([^_]+)_(.*?)_\[(.*?)\]_\(\d{6}\)$"
        ma = re.match(pattern_a, name_no_ext)
        if ma:
            seq, c_type, prop_info, parties = ma.groups()
            new_f = f"{seq}_{c_type}_{prop_info}_[{parties}]_({final_date_str}).pdf"
        else:
            new_f = f

        date_proposals.append({
            "folder": f"[{parent_name}] {folder_name}",
            "old": f,
            "new": new_f,
            "lease_start": final_date_str
        })

print(f"=== LEASE START DATE RENAMING PROPOSAL (TOTAL {len(date_proposals)} FILES) ===")
for item in date_proposals[:15]:
    print(f"\n폴더: {item['folder']}")
    print(f"  - 기존 체결일자 파일명: {item['old']}")
    print(f"  - 변경 임대시작일 파일명: {item['new']}")
