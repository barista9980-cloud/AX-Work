import os
import fitz # PyMuPDF
from rapidocr_onnxruntime import RapidOCR
import re

ocr_engine = RapidOCR()

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

pdf_files = []
for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    if "402_403" in root or "1510" in root:
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
        
        # OCR page 1
        page = doc[0]
        pix = page.get_pixmap(dpi=200)
        img_bytes = pix.tobytes("png")
        
        result, _ = ocr_engine(img_bytes)
        
        print("--- OCR RESULTS PAGE 1 ---")
        ocr_lines = []
        if result:
            for item in result:
                txt = item[1]
                score = item[2]
                ocr_lines.append(txt)
                print(f"[{score:.2f}] {txt}")
                
        full_text = "\n".join(ocr_lines)
        
        # Search for lease period patterns
        # Look for dates like 20XX년 XX월 XX일 ~ 20XX년 XX월 XX일 or XX개월
        date_patterns = [
            r"(\d{4}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일)\s*(?:부터|~|-)\s*(\d{4}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일)",
            r"(\d{4}\.\s*\d{1,2}\.\s*\d{1,2})\s*(?:~|-)\s*(\d{4}\.\s*\d{1,2}\.\s*\d{1,2})",
            r"(\d{2,4}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일)",
            r"보증금.*?\d+",
            r"임대료.*?\d+",
            r"차임.*?\d+"
        ]
        
        print("\n--- SEARCHED PATTERNS ---")
        for pat in date_patterns:
            matches = re.findall(pat, full_text)
            if matches:
                print(f"Pattern '{pat}': {matches}")
                
    except Exception as e:
        print(f"Error in OCR: {e}")
