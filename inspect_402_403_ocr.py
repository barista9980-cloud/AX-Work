import sys
import fitz
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001\가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229).pdf"

doc = fitz.open(pdf_path)
print(f"Total pages in 402_403호 PDF: {len(doc)}")

engine = RapidOCR()

for page_num in range(len(doc)):
    print(f"\n==================== Page {page_num + 1} ====================")
    pix = doc[page_num].get_pixmap(dpi=250)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    results, _ = engine(np.array(img))
    if results:
        for idx, item in enumerate(results, 1):
            text = item[1]
            score = item[2]
            print(f"L{idx:02d}: {text}")
    else:
        print("No text detected.")
