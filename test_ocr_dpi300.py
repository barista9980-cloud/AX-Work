import os
import fitz # PyMuPDF
from rapidocr_onnxruntime import RapidOCR
import cv2
import numpy as np

ocr_engine = RapidOCR()

target_pdf = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차\01_가산_대륭포스트타워6차\402_403호\가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229).pdf"

doc = fitz.open(target_pdf)
page = doc[0]

# Render at 300 DPI
pix = page.get_pixmap(dpi=300)
img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))

if pix.n == 4: # RGBA
    img_np = cv2.cvtColor(img_np, cv2.COLOR_RGBA2BGR)

# Run RapidOCR on 300 DPI image
result, _ = ocr_engine(img_np)

print("=== 300 DPI OCR RESULTS ===")
if result:
    for idx, item in enumerate(result):
        txt = item[1]
        score = item[2]
        print(f"[{idx:02d}] {txt}")
