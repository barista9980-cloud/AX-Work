import os
import re
import json
import fitz
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

DOWNLOAD_DIR = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"
JSON_PATH = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\real_estate_parsed_catalog.json"

ocr_engine = RapidOCR()

KNOWN_BANKS = ["국민은행", "신한은행", "하나은행", "우리은행", "기업은행", "농협", "새마을금고", "부산은행", "대구은행", "카카오뱅크"]

def extract_strict_clean_info(pdf_path):
    deposit, rent, period, bank, account, phone, area = "", "", "", "", "", "", ""
    try:
        doc = fitz.open(pdf_path)
        ocr_texts = []
        for p_idx in range(min(2, len(doc))):  # Main 1-2 contract pages only
            pix = doc[p_idx].get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            results, _ = ocr_engine(np.array(img))
            if results:
                for res in results:
                    ocr_texts.append(res[1])

        combined_text = "\n".join(ocr_texts)

        # 1. Bank
        for b in KNOWN_BANKS:
            if b in combined_text or b.replace("은행", "") in combined_text:
                bank = b
                break

        # 2. Account Number: Strict regex (must be near 계좌 or 입금 or 은행, or match clear bank pattern)
        acc_m = re.search(r'(?:계좌|입금|농협|국민|신한|우리|하나|기업)[^\d\n]*(\d{3,6}[\-\s]\d{2,6}[\-\s]\d{3,7})', combined_text)
        if not acc_m:
            # Fallback: standalone account pattern, but exclude dates/business numbers/phone prefixes
            acc_candidates = re.findall(r'(\d{3,6}\-\d{2,6}\-\d{3,7})', combined_text)
            for cand in acc_candidates:
                if not cand.startswith(("010-", "011-", "02-", "031-", "070-", "106-", "697-", "2026", "2025", "2024", "2023", "2022")):
                    account = cand
                    break
        else:
            account = acc_m.group(1).strip()

        if account and not bank:
            if account.startswith("415"):
                bank = "국민은행"
            elif account.startswith("110"):
                bank = "신한은행"

        # 3. Phone: Extract 1-2 valid phone numbers (010, 02, 031, 070)
        phone_matches = re.findall(r'(010[\-\s]\d{4}[\-\s]\d{4}|02[\-\s]\d{3,4}[\-\s]\d{4}|031[\-\s]\d{3,4}[\-\s]\d{4}|070[\-\s]\d{3,4}[\-\s]\d{4})', combined_text)
        if phone_matches:
            phone = ", ".join(list(set(phone_matches))[:2])

        # 4. Deposit & Rent
        # Deposit: 보증금 XXX 원 or 전세금 XXX 원
        dep_m = re.search(r'(?:보증금|전세금)[^\d\n]*([\d\,]{5,12})\s*원', combined_text)
        if dep_m:
            deposit = dep_m.group(1).strip() + " 원"

        rent_m = re.search(r'(?:차임|월임대료|월세|월\s*대여료)[^\d\n]*([\d\,]{4,10})\s*원', combined_text)
        if rent_m:
            rent = rent_m.group(1).strip() + " 원"

        # 5. Lease Period
        per_m = re.search(r'(20\d{2}[\.\-\s년]\s*\d{1,2}[\.\-\s월]\s*\d{1,2})[^\d\n\~]*\~[^\d\n]*(20\d{2}[\.\-\s년]\s*\d{1,2}[\.\-\s월]\s*\d{1,2})', combined_text)
        if per_m:
            period = f"{per_m.group(1).strip()} ~ {per_m.group(2).strip()}"

        # 6. Area
        area_m = re.search(r'(?:면적|계약면적|전용면적)[^\d\n]*([\d\.]+)\s*(㎡|제곱미터|평)', combined_text)
        if area_m:
            area = f"{area_m.group(1)} {area_m.group(2)}"

    except Exception as e:
        pass

    return {
        "deposit": deposit,
        "rent": rent,
        "period": period,
        "bank": bank,
        "account": account,
        "phone": phone,
        "area": area
    }

def main():
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        catalog = json.load(f)

    print("--- Strict Clean OCR Inspection Test ---")
    for idx, item in enumerate(catalog[:10], 1):
        fname = item["filename"]
        path = os.path.join(DOWNLOAD_DIR, fname)
        info = extract_strict_clean_info(path)
        print(f"[{idx:02d}] {fname[:35]}...\n     Dep: '{info['deposit']}', Rent: '{info['rent']}', Bank: '{info['bank']}', Acc: '{info['account']}', Phone: '{info['phone']}'")

if __name__ == "__main__":
    main()
