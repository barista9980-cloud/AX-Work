import os
import fitz # PyMuPDF
from rapidocr_onnxruntime import RapidOCR

ocr_engine = RapidOCR()

target_pdf = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차\01_가산_대륭포스트타워6차\402_403호\가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229).pdf"

if os.path.exists(target_pdf):
    doc = fitz.open(target_pdf)
    pix = doc[0].get_pixmap(dpi=200)
    img_bytes = pix.tobytes("png")
    result, _ = ocr_engine(img_bytes)
    print("=== OCR RESULTS FOR 402_403호 PAGE 1 ===")
    if result:
        for idx, item in enumerate(result):
            print(f"[{idx:02d}] {item[1]}")
else:
    print("PDF not found!")
