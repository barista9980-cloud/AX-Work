import os
import sys
import fitz # PyMuPDF
from rapidocr_onnxruntime import RapidOCR
import re

sys.stdout.reconfigure(encoding='utf-8')
ocr_engine = RapidOCR()

fox_root = r"G:\내 드라이브\[FoxConnect]"

pdf_files = []
for root, dirs, files in os.walk(fox_root):
    if "402_403" in root:
        for f in files:
            if f.endswith(".pdf"):
                pdf_files.append(os.path.join(root, f))

out_lines = []

for pdf_p in pdf_files:
    out_lines.append(f"\n==========================================")
    out_lines.append(f"OCR Parsing Scanned PDF: {os.path.basename(pdf_p)}")
    out_lines.append(f"==========================================")
    try:
        doc = fitz.open(pdf_p)
        out_lines.append(f"Total Pages: {len(doc)}")
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=200)
            img_bytes = pix.tobytes("png")
            
            result, _ = ocr_engine(img_bytes)
            
            out_lines.append(f"\n--- OCR RESULTS PAGE {page_num+1} ---")
            ocr_text_list = []
            if result:
                for item in result:
                    txt = item[1]
                    ocr_text_list.append(txt)
                    out_lines.append(txt)
                    
            full_text = "\n".join(ocr_text_list)
            
            out_lines.append("\n--- KEYWORD / DATE SEARCH ---")
            for line in ocr_text_list:
                if any(kw in line for kw in ["임대기간", "계약기간", "기간", "년", "월", "일", "부터", "까지", "보증금", "월세", "임대료"]):
                    out_lines.append(f"  [MATCH] {line}")
                    
    except Exception as e:
        out_lines.append(f"Error in OCR: {e}")

with open("ocr_results_402_403.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out_lines))

print("Saved OCR results to ocr_results_402_403.txt")
