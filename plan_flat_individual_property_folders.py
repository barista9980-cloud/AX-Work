import os
import sys
import re
import json

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
REAL_ESTATE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리")

units_list = []

for root, dirs, files in os.walk(REAL_ESTATE_BASE):
    pdf_files = [f for f in files if f.lower().endswith(".pdf")]
    if not pdf_files:
        continue

    # Find earliest lease start date among PDF files in this unit folder
    earliest_date = "9999-99-99"
    for f in pdf_files:
        name_no_ext = os.path.splitext(f)[0]
        m_date = re.search(r"\((\d{6})\)", name_no_ext)
        if m_date:
            raw_d = m_date.group(1)
            iso_d = f"20{raw_d[:2]}-{raw_d[2:4]}-{raw_d[4:6]}"
            if iso_d < earliest_date:
                earliest_date = iso_d

    folder_name = os.path.basename(root)
    parent_name = os.path.basename(os.path.dirname(root))

    # Clean building and unit names
    clean_parent = re.sub(r"^\d{2}_", "", parent_name)
    clean_folder = re.sub(r"^\d{2}_", "", folder_name)

    if clean_parent in ["01_임대차계약", "02_매매_소유권문서", "01_부동산_자산관리"]:
        combined_label = clean_folder
    else:
        combined_label = f"{clean_parent}_{clean_folder}"

    units_list.append({
        "current_path": root,
        "combined_label": combined_label,
        "earliest_date": earliest_date if earliest_date != "9999-99-99" else "미상",
        "pdf_count": len(pdf_files)
    })

units_list.sort(key=lambda x: x["earliest_date"])

print(f"=== FLAT INDIVIDUAL PROPERTY FOLDERS PLAN (TOTAL {len(units_list)} UNITS) ===")
flat_plan = []
for idx, u in enumerate(units_list, 1):
    seq_str = f"{idx:02d}"
    flat_folder_name = f"{seq_str}_{u['combined_label']}"
    flat_plan.append({
        "seq": seq_str,
        "new_folder_name": flat_folder_name,
        "earliest_date": u["earliest_date"],
        "pdf_count": u["pdf_count"],
        "current_path": u["current_path"]
    })
    print(f"[{seq_str}] {flat_folder_name} (최초계약일: {u['earliest_date']}, 계약서 {u['pdf_count']}개)")

with open("flat_folder_plan.json", "w", encoding="utf-8") as f:
    json.dump(flat_plan, f, ensure_ascii=False, indent=2)
