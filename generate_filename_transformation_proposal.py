import os
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

proposals = []

for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    pdf_files = [f for f in files if f.lower().endswith(".pdf")]
    if not pdf_files:
        continue

    folder_name = os.path.basename(root)
    parent_name = os.path.basename(os.path.dirname(root))

    pdf_files.sort()
    for idx, f in enumerate(pdf_files, 1):
        name_no_ext = os.path.splitext(f)[0]
        
        # Parse current filename components
        pattern = r"^(.*?)_(\d{2})_([^_]+)_\[(.*?)\]_\((\d{6})\)$"
        m = re.match(pattern, name_no_ext)
        if m:
            prop_info, seq, c_type, parties, raw_date = m.groups()
            
            # Proposed Option A: 01_최초임대차_건물_호수_[당사자]_(YYMMDD).pdf
            new_opt_a = f"{seq}_{c_type}_{prop_info}_[{parties}]_({raw_date}).pdf"
            
            # Proposed Option B: 01_최초임대차_[당사자]_(YYMMDD).pdf (Clean & concise inside folder)
            new_opt_b = f"{seq}_{c_type}_[{parties}]_({raw_date}).pdf"
            
            proposals.append({
                "folder": f"[{parent_name}] {folder_name}",
                "old": f,
                "opt_a": new_opt_a,
                "opt_b": new_opt_b
            })

print("=== FILENAME TRANSFORMATION PROPOSALS (SAMPLE 15 FILES) ===")
for p in proposals[:15]:
    print(f"\nFolder: {p['folder']}")
    print(f"  [현재 파일명] : {p['old']}")
    print(f"  [추천 A형식] : {p['opt_a']}  (순서맨앞+구분+건물+당사자+일자)")
    print(f"  [추천 B형식] : {p['opt_b']}  (순서맨앞+구분+당사자+일자 - 폴더내 간결형)")
