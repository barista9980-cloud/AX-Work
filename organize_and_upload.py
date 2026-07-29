import os
import shutil
import json

BASE_DIR = os.path.dirname(__file__)
DOWNLOAD_DIR = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"
TARGET_ROOT = os.path.join(BASE_DIR, "gdrive_structure", "[부동산자산] FoxConnect 계약 관리")
CATALOG_PATH = os.path.join(BASE_DIR, "real_estate_parsed_catalog.json")

REGION_MAP = {
    "대전": "01_대전",
    "광명": "02_광명",
    "강남": "03_강남",
    "가산": "04_가산",
    "판교": "05_판교",
    "서초": "06_서초",
    "세종": "07_세종"
}

def organize_local_structure():
    if not os.path.exists(CATALOG_PATH):
        print("Catalog JSON not found.")
        return

    with open(CATALOG_PATH, "r", encoding="utf-8") as f:
        catalog = json.load(f)

    # Create root and subdirectories
    subdirs = [
        os.path.join(TARGET_ROOT, "01_권역별_계약서"),
        os.path.join(TARGET_ROOT, "02_서식_템플릿"),
        os.path.join(TARGET_ROOT, "03_자동파싱_업로드"),
        os.path.join(TARGET_ROOT, "04_생성_보고서")
    ]
    for d in subdirs:
        os.makedirs(d, exist_ok=True)

    # Create region folders
    for reg_code in REGION_MAP.values():
        os.makedirs(os.path.join(TARGET_ROOT, "01_권역별_계약서", reg_code), exist_ok=True)

    # Copy files into region folders
    copied_count = 0
    for item in catalog:
        filename = item["filename"]
        region = item.get("region", "기타")
        src_path = os.path.join(DOWNLOAD_DIR, filename)
        
        reg_folder = REGION_MAP.get(region, "08_기타")
        dest_dir = os.path.join(TARGET_ROOT, "01_권역별_계약서", reg_folder)
        os.makedirs(dest_dir, exist_ok=True)
        
        dest_path = os.path.join(dest_dir, filename)
        if os.path.exists(src_path):
            shutil.copy2(src_path, dest_path)
            copied_count += 1

    # Copy master CSV ledger into root
    csv_src = os.path.join(BASE_DIR, "FoxConnect_부동산_자산대장_1차목록.csv")
    if os.path.exists(csv_src):
        shutil.copy2(csv_src, os.path.join(TARGET_ROOT, "FoxConnect_부동산_자산대장_1차목록.csv"))

    print(f"Organized {copied_count} files into structured directory: {TARGET_ROOT}")

if __name__ == "__main__":
    organize_local_structure()
