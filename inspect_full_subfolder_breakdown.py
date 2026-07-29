import os
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
LEASE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리\01_임대차계약")

building_folders = []
for item in os.listdir(LEASE_BASE):
    full_p = os.path.join(LEASE_BASE, item)
    if os.path.isdir(full_p):
        building_folders.append((item, full_p))

folder_earliest_dates = []

for item, full_p in building_folders:
    earliest_date = "9999-99-99"
    subunits = []

    for sub in os.listdir(full_p):
        sub_p = os.path.join(full_p, sub)
        if os.path.isdir(sub_p):
            pdf_files = [f for f in os.listdir(sub_p) if f.lower().endswith(".pdf")]
            unit_date = "9999-99-99"
            for f in pdf_files:
                name_no_ext = os.path.splitext(f)[0]
                m_date = re.search(r"\((\d{6})\)", name_no_ext)
                if m_date:
                    raw_d = m_date.group(1)
                    iso_d = f"20{raw_d[:2]}-{raw_d[2:4]}-{raw_d[4:6]}"
                    if iso_d < unit_date:
                        unit_date = iso_d
                    if iso_d < earliest_date:
                        earliest_date = iso_d
            subunits.append((sub, unit_date if unit_date != "9999-99-99" else "미상", len(pdf_files)))

    clean_name = re.sub(r"^\d{2}_", "", item)
    folder_earliest_dates.append({
        "old_folder_name": item,
        "clean_name": clean_name,
        "full_path": full_p,
        "earliest_date": earliest_date if earliest_date != "9999-99-99" else "미상",
        "subunits": subunits
    })

folder_earliest_dates.sort(key=lambda x: x["earliest_date"])

print("=== FULL BUILDING & UNIT SUBFOLDER BREAKDOWN ===")
for idx, b in enumerate(folder_earliest_dates, 1):
    seq_str = f"{idx:02d}"
    print(f"\n[{seq_str}] 건물: {seq_str}_{b['clean_name']} (최초계약: {b['earliest_date']})")
    for u_name, u_date, p_cnt in b["subunits"]:
        print(f"    └─ 🏠 호수/구역: {u_name} (계약서 {p_cnt}개, 시작일: {u_date})")
