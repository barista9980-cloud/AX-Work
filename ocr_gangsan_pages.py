import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    import pytesseract
    from PIL import Image

    # Set tesseract cmd path if needed on Windows
    tesseract_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Users\User\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
    ]
    for p in tesseract_paths:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            break

    print("Running OCR on Gangsan rendered pages...")
    img_dir = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\gangsan_pages"
    for img_name in sorted(os.listdir(img_dir)):
        if img_name.endswith(".png"):
            ip = os.path.join(img_dir, img_name)
            img = Image.open(ip)
            txt = pytesseract.image_to_string(img, lang='kor+eng')
            print(f"\n--- {img_name} ---")
            for line in txt.split("\n"):
                if any(k in line for k in ["보증금", "금", "원", "삼억", "오억", "십억", "300", "500", "000", "차임", "월세", "임대료", "박재윤", "강산"]):
                    print("  ", line.strip())
except Exception as e:
    print("OCR Error / Not Installed:", e)
