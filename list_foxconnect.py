import os

foxconnect_path = r"G:\내 드라이브\[FoxConnect]"

print(f"Listing contents of {foxconnect_path}:")

for root, dirs, files in os.walk(foxconnect_path):
    print(f"\nDirectory: {root}")
    for d in dirs:
        print(f"  [DIR] {d}")
    for f in files:
        print(f"  [FILE] {f}")
