import os
import sys
import re
import json

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
LEASE_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리\01_임대차계약")

print("Analyzing building folders in:", LEASE_BASE)

building_folders = []
for item in os.listdir(LEASE_BASE):
    full_p = os.path.join(LEASE_BASE, item)
    if os.path.isdir(full_p):
        building_folders.append((item, full_p))

folder_earliest_dates = []

for item, full_p in building_folders:
    # Find all pdf files inside this building folder and its subfolders
    earliest_date = "9999-99-99"
    pdf_count = 0

    for root, dirs, files in os.walk(full_p):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdf_count += 1
                name_no_ext = os.path.splitext(f)[0]
                m_date = re.search(r"\((\d{6})\)", name_no_ext)
                if m_date:
                    raw_d = m_date.group(1)
                    y = "20" + raw_d[:2]
                    mon = raw_d[2:4]
                    d = raw_d[4:6]
                    iso_d = f"{y}-{mon}-{d}"
                    if iso_d < earliest_date:
                        earliest_date = iso_d

    # Strip existing prefix like "01_", "02_"
    clean_name = re.sub(r"^\d{2}_", "", item)
    folder_earliest_dates.append({
        "old_folder_name": item,
        "clean_name": clean_name,
        "full_path": full_p,
        "earliest_date": earliest_date if earliest_date != "9999-99-99" else "미상",
        "pdf_count": pdf_count
    })

# Sort building folders chronologically by earliest lease start date
folder_earliest_dates.sort(key=lambda x: x["earliest_date"])

print(f"Total Building Folders Analyzed: {len(folder_earliest_dates)}")

proposed_folder_plan = []
for idx, f_info in enumerate(folder_earliest_dates, 1):
    seq_str = f"{idx:02d}"
    new_folder_name = f"{seq_str}_{f_info['clean_name']}"
    proposed_folder_plan.append({
        "seq": seq_str,
        "old": f_info["old_folder_name"],
        "new": new_folder_name,
        "earliest_date": f_info["earliest_date"],
        "full_path": f_info["full_path"]
    })

print("\n=== PROPOSED CHRONOLOGICAL BUILDING FOLDER RE-ORDERING PLAN ===")
for p in proposed_folder_plan:
    print(f"[{p['seq']}] {p['new']} (최초계약시작일: {p['earliest_date']}) <- 기존: {p['old']}")

with open("proposed_folder_plan.json", "w", encoding="utf-8") as f:
    json.dump(proposed_folder_plan, f, ensure_ascii=False, indent=2)
