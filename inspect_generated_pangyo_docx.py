import os
import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\부동산_계약관리노트_판교_판교동612.docx"

print("Inspecting generated docx:", docx_path)
doc = docx.Document(docx_path)

for t_idx, table in enumerate(doc.tables):
    print(f"\n--- TABLE {t_idx} ---")
    for r_idx, row in enumerate(table.rows):
        cells_txt = [c.text.replace("\n", " ") for c in row.cells]
        print(f" Row {r_idx}:", " | ".join(cells_txt))
