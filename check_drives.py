import os
import sys

drives = ["C:", "D:", "E:", "F:", "G:", "H:", "Z:"]

for d in drives:
    if os.path.exists(d + "\\"):
        print(f"Drive {d} exists.")
        try:
            items = os.listdir(d + "\\")
            print(f"  Root items in {d}: {items[:10]}")
        except Exception as e:
            print(f"  Cannot list {d}: {e}")

gdrive_path = r"G:\내 드라이브"
if os.path.exists(gdrive_path):
    print("Found G:\\내 드라이브:", os.listdir(gdrive_path))

download_sample_dir = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"
if os.path.exists(download_sample_dir):
    print("Found Downloads sample dir with", len(os.listdir(download_sample_dir)), "files.")
