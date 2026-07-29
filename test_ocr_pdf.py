import sys
import fitz
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001\가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229).pdf"

doc = fitz.open(pdf_path)
print(f"Total pages: {len(doc)}")

engine = RapidOCR()

page = doc[0]
pix = page.get_pixmap(dpi=200)
img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

img_np = np.array(img)
results, elapse = engine(img_np)

print(f"\n--- OCR Extracted Text (Page 1) ---")
if results:
    for item in results:
        text = item[1]
        score = item[2]
        print(f"[{score}] {text}")
else:
        print("No text detected.")
