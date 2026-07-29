import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

pdf_by_folder = {}

for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    pdfs = [f for f in files if f.lower().endswith(".pdf")]
    if pdfs:
        rel_path = os.path.relpath(root, FOXCONNECT_ROOT)
        pdf_by_folder[rel_path] = pdfs

total_pdfs = sum([len(v) for v in pdf_by_folder.values()])

print(f"=== TOTAL PDF CONTRACT FILES COUNT: {total_pdfs} ===")
for folder, pdf_list in pdf_by_folder.items():
    print(f"\n📁 [{folder}] - {len(pdf_list)}개")
    for f in pdf_list:
        print(f"   - {f}")
