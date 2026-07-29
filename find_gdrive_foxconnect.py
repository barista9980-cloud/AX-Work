import os
import glob

possible_roots = [
    r"G:\내 드라이브\[FoxConnect]",
    r"G:\My Drive\[FoxConnect]",
    r"C:\Users\User\Google Drive\[FoxConnect]",
    r"C:\Users\User\OneDrive\[FoxConnect]",
    r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"
]

print("Checking possible [FoxConnect] paths...")

found = []

for p in possible_roots:
    if os.path.exists(p):
        found.append(p)
        print(f"[FOUND] {p}")
    else:
        print(f"[NOT FOUND] {p}")

# Search drive letters
for letter in "DEFG":
    drive_path = f"{letter}:\\"
    if os.path.exists(drive_path):
        print(f"Checking drive {drive_path}...")
        try:
            for root, dirs, files in os.walk(drive_path):
                if "[FoxConnect]" in root:
                    print(f"[WALK FOUND] {root}")
                    found.append(root)
                    break
                # Only check top 2 levels
                if root.count(os.sep) > 3:
                    dirs.clear()
        except Exception as e:
            pass

print("Search complete.")
