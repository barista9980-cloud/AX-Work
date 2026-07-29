import os
import sys
import fitz # PyMuPDF
from rapidocr_onnxruntime import RapidOCR
import cv2
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
ocr_engine = RapidOCR()

fox_root = r"G:\내 드라이브\[FoxConnect]"

pdf_path = None
for root, dirs, files in os.walk(fox_root):
    if "402_403" in root:
        for f in files:
            if f.endswith(".pdf") and "01_최초임대차" in f:
                pdf_path = os.path.join(root, f)
                break
    if pdf_path:
        break

print(f"Target PDF Path: {pdf_path}")

if pdf_path:
    doc = fitz.open(pdf_path)
    print(f"Total Pages in PDF: {len(doc)}")
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=300)
        img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))
        
        if pix.n == 4:
            img_np = cv2.cvtColor(img_np, cv2.COLOR_RGBA2BGR)
        elif pix.n == 1:
            img_np = cv2.cvtColor(img_np, cv2.COLOR_GRAY2BGR)
            
        result, _ = ocr_engine(img_np)
        
        print(f"\n=== PAGE {page_num+1} 300 DPI OCR RESULTS ===")
        if result:
            for idx, item in enumerate(result):
                print(f"[{idx:02d}] {item[1]}")
