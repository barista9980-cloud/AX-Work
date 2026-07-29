"""
Universal Folder Structure & Upload Queue Initialization Engine
"""
import os
import sys

def init_corporate_folder_structure(base_dir):
    """
    Initializes standard corporate asset management directories and upload queues.
    """
    print(f"[FolderEngine] Initializing Corporate Asset Directories in: {base_dir}")

    structure = {
        "01_부동산_자산관리": [
            "00_연도별_부동산_총괄자산대장",
            "01_임대차계약",
            "02_매매_소유권문서"
        ],
        "02_차량_자산관리": [
            "00_연도별_차량_총괄자산대장",
            "01_차량계약_리스_렌트"
        ],
        "03_보험_자산관리": [
            "00_연도별_보험_총괄자산대장",
            "01_보험증권_및_배서계약서",
            "02_보험금청구_사고접수",
            "03_보험료납입_증빙"
        ],
        "06_자동파싱_업로드큐": [
            "01_부동산_업로드대기",
            "02_차량_업로드대기",
            "03_보험_업로드대기",
            "04_비품_소모품_업로드대기",
            "05_처리완료_아카이브"
        ]
    }

    for main_dir, sub_dirs in structure.items():
        main_path = os.path.join(base_dir, main_dir)
        os.makedirs(main_path, exist_ok=True)
        print(f"  + {main_dir}/")
        for sub_d in sub_dirs:
            sub_path = os.path.join(main_path, sub_d)
            os.makedirs(sub_path, exist_ok=True)
            print(f"    └── {sub_d}/")

    print("[FolderEngine] Directory Initialization Complete!\n")

if __name__ == "__main__":
    from src.config import DEFAULT_BASE_DIR
    init_corporate_folder_structure(DEFAULT_BASE_DIR)
