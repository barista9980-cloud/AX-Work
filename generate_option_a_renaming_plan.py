import os
import sys
import re
import json

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

renaming_plan = []

for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    pdf_files = [f for f in files if f.lower().endswith(".pdf")]
    if not pdf_files:
        continue

    folder_name = os.path.basename(root)
    parent_name = os.path.basename(os.path.dirname(root))

    # Sort pdf files logically: 최초임대차 first, then by date
    def sort_key(f_name):
        name_no_ext = os.path.splitext(f_name)[0]
        is_first = 0 if "최초임대차" in name_no_ext or "매매" in name_no_ext else 1
        date_match = re.search(r"\((\d{6})\)", name_no_ext)
        d_val = date_match.group(1) if date_match else "999999"
        return (is_first, d_val, f_name)

    pdf_files.sort(key=sort_key)

    for idx, f in enumerate(pdf_files, 1):
        name_no_ext = os.path.splitext(f)[0]
        seq_str = f"{idx:02d}"

        pattern = r"^(.*?)_(\d{2})_([^_]+)_\[(.*?)\]_\((\d{6})\)$"
        m = re.match(pattern, name_no_ext)
        if m:
            prop_info, old_seq, c_type, parties, raw_date = m.groups()
            new_filename = f"{seq_str}_{c_type}_{prop_info}_[{parties}]_({raw_date}).pdf"
        else:
            # Fallback parsing
            date_match = re.search(r"\((\d{6})\)", name_no_ext)
            raw_date = date_match.group(1) if date_match else "240101"
            
            parties_match = re.search(r"\[(.*?)\]", name_no_ext)
            parties = parties_match.group(1) if parties_match else ""
            
            c_type = "최초임대차"
            if "전대차" in name_no_ext:
                c_type = "전대차"
            elif "변경계약" in name_no_ext:
                c_type = "변경계약"
            elif "연장" in name_no_ext:
                c_type = "연장계약"
            elif "매매" in name_no_ext:
                c_type = "매매"
                
            new_filename = f"{seq_str}_{c_type}_{folder_name}_[{parties}]_({raw_date}).pdf"

        renaming_plan.append({
            "folder_path": root,
            "folder_label": f"[{parent_name}] {folder_name}",
            "old_filename": f,
            "new_filename": new_filename
        })

print(f"Total PDF Files to Rename: {len(renaming_plan)}")

with open("renaming_plan_option_a.json", "w", encoding="utf-8") as f:
    json.dump(renaming_plan, f, ensure_ascii=False, indent=2)

print("\n--- SAMPLE RENAMING PLAN (FIRST 20 FILES) ---")
for item in renaming_plan[:20]:
    print(f"\n폴더: {item['folder_label']}")
    print(f"  - 기존: {item['old_filename']}")
    print(f"  - 변경: {item['new_filename']}")
