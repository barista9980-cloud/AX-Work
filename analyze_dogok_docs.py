import os
import re
import fitz  # PyMuPDF
import json
from PIL import Image
import numpy as np
from rapidocr_onnxruntime import RapidOCR

DOWNLOAD_DIR = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"
ocr_engine = RapidOCR()

target_files = [
    f for f in os.listdir(DOWNLOAD_DIR)
    if "도곡로1길23" in f and f.endswith(".pdf")
]

target_files.sort()

results = []

for filename in target_files:
    filepath = os.path.join(DOWNLOAD_DIR, filename)
    doc = fitz.open(filepath)
    
    full_text = ""
    for page in doc:
        text = page.get_text()
        if text and len(text.strip()) > 50:
            full_text += text + "\n"
        else:
            # OCR fallback
            pix = page.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            res, _ = ocr_engine(np.array(img))
            if res:
                full_text += " ".join([r[1] for r in res]) + "\n"
                
    results.append({
        "filename": filename,
        "page_count": len(doc),
        "text_length": len(full_text),
        "full_text": full_text
    })

output_json = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\dogok_extracted_text.json"
with open(output_json, "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"Extracted text from {len(results)} files.")
