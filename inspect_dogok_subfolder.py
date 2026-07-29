import os
import sys
import docx
import fitz

sys.stdout.reconfigure(encoding='utf-8')

target_dir = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\02_강남_도곡로1길23\지하1층,1층,2층,3층"

files = os.listdir(target_dir)
print("Files in Dogok 지하1층,1층,2층,3층 folder:")
for f in files:
    print(" -", f)

# Check generated docx note
docx_file = [f for f in files if f.endswith(".docx") and "계약관리노트" in f]
if docx_file:
    docx_path = os.path.join(target_dir, docx_file[0])
    print(f"\n=== CURRENT DOCX NOTE ({docx_file[0]}) ===")
    doc = docx.Document(docx_path)
    for t_idx, t in enumerate(doc.tables):
        print(f"\n--- TABLE {t_idx} ---")
        for r_idx, r in enumerate(t.rows):
            row_str = " | ".join([c.text.strip().replace('\n', ' // ') for c in r.cells])
            print(f"  Row {r_idx}: {row_str}")

# Inspect PDFs
pdf_files = [f for f in files if f.endswith(".pdf")]
print(f"\n=== PDF FILES IN DOGOK ({len(pdf_files)} files) ===")
for pdf_f in pdf_files:
    pdf_path = os.path.join(target_dir, pdf_f)
    doc_pdf = fitz.open(pdf_path)
    print(f"\nFile: {pdf_f} (Pages: {len(doc_pdf)})")
    for p_idx in range(len(doc_pdf)):
        page_text = doc_pdf[p_idx].get_text()
        print(f"  --- Page {p_idx+1} Text Snippet ---\n", page_text[:500])
