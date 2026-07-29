import os
import fitz # PyMuPDF
import re

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

pdf_files = []
for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    if "402_403" in root or "1510" in root:
        for f in files:
            if f.endswith(".pdf"):
                pdf_files.append(os.path.join(root, f))

for pdf_p in pdf_files:
    print(f"\n==========================================")
    print(f"Parsing PDF: {os.path.basename(pdf_p)}")
    print(f"==========================================")
    try:
        doc = fitz.open(pdf_p)
        print(f"Total Pages: {len(doc)}")
        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text")
            print(f"\n--- PAGE {page_num+1} TEXT ---")
            lines = [line.strip() for line in text.split("\n") if line.strip()]
            for l in lines:
                print(l)
    except Exception as e:
        print(f"Error parsing PDF: {e}")
