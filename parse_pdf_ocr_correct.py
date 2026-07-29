import os
import fitz # PyMuPDF
from rapidocr_onnxruntime import RapidOCR
import re

ocr_engine = RapidOCR()

fox_root = r"G:\내 드라이브\[FoxConnect]"

pdf_files = []
for root, dirs, files in os.walk(fox_root):
    if "402_403" in root:
        for f in files:
            if f.endswith(".pdf"):
                pdf_files.append(os.path.join(root, f))

for pdf_p in pdf_files:
    print(f"\n==========================================")
    print(f"OCR Parsing Scanned PDF: {os.path.basename(pdf_p)}")
    print(f"==========================================")
    try:
        doc = fitz.open(pdf_p)
        print(f"Total Pages: {len(doc)}")
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=200)
            img_bytes = pix.tobytes("png")
            
            result, _ = ocr_engine(img_bytes)
            
            print(f"\n--- OCR RESULTS PAGE {page_num+1} ---")
            ocr_lines = []
            if result:
                for item in result:
                    txt = item[1]
                    ocr_lines.append(txt)
                    print(f"{txt}")
                    
            full_text = "\n".join(ocr_lines)
            
            print("\n--- DATES AND PERIOD SEARCH ---")
            # Search for all date patterns
            dates_found = re.findall(r"\d{4}년\s*\d{1,2}월\s*\d{1,2}일|\d{4}\.\s*\d{1,2}\.\s*\d{1,2}|\d{2}년\s*\d{1,2}월\s*\d{1,2}일", full_text)
            print("Dates found:", dates_found)
            
            # Search for period lines
            for line in ocr_lines:
                if any(kw in line for kw in ["임대기간", "계약기간", "기간", "년", "월", "일", "부터", "까지", "보증금", "월세", "임대료"]):
                    print(f"  [KEYWORD MATCH] {line}")
                    
    except Exception as e:
        print(f"Error in OCR: {e}")
