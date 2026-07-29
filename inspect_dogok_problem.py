import os
import sys
import docx
import fitz
from google import genai
from google.genai import types
import json

sys.stdout.reconfigure(encoding='utf-8')

USER_API_KEY = os.environ.get("GEMINI_API_KEY", "")
client = genai.Client(api_key=USER_API_KEY) if USER_API_KEY else None

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

dogok_dir = None
for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    if "도곡로1길23" in root:
        dogok_dir = root
        break

print("Dogok directory found:", dogok_dir)

if dogok_dir:
    files = os.listdir(dogok_dir)
    print("\nFiles in Dogok folder:")
    for f in files:
        print(" -", f)

    # Check generated docx note
    docx_file = [f for f in files if f.endswith(".docx") and "계약관리노트" in f]
    if docx_file:
        docx_path = os.path.join(dogok_dir, docx_file[0])
        print(f"\n=== CURRENT DOCX NOTE ({docx_file[0]}) ===")
        doc = docx.Document(docx_path)
        for t_idx, t in enumerate(doc.tables):
            print(f"\n--- TABLE {t_idx} ---")
            for r_idx, r in enumerate(t.rows):
                row_str = " | ".join([c.text.strip().replace('\n', ' // ') for c in r.cells])
                print(f"  Row {r_idx}: {row_str}")

    # Inspect PDFs with PyMuPDF / Vision AI
    pdf_files = [f for f in files if f.endswith(".pdf")]
    print(f"\n=== PDF FILES IN DOGOK ({len(pdf_files)} files) ===")
    for pdf_f in pdf_files:
        pdf_path = os.path.join(dogok_dir, pdf_f)
        doc_pdf = fitz.open(pdf_path)
        print(f"\nFile: {pdf_f} (Pages: {len(doc_pdf)})")
        page_text = doc_pdf[0].get_text()
        print("  Text snippet from page 1:\n", page_text[:400])
