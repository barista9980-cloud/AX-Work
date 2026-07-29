import os
import shutil

SRC_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\gdrive_structure\[부동산자산] FoxConnect 계약 관리"
GDRIVE_DEST = r"G:\내 드라이브\[부동산자산] FoxConnect 계약 관리"

def sync_to_google_drive():
    if not os.path.exists(SRC_DIR):
        print(f"Source directory not found: {SRC_DIR}")
        return

    print(f"Syncing from {SRC_DIR} to {GDRIVE_DEST}...")
    shutil.copytree(SRC_DIR, GDRIVE_DEST, dirs_exist_ok=True)
    print(f"Successfully synced to Google Drive: {GDRIVE_DEST}")

if __name__ == "__main__":
    sync_to_google_drive()
