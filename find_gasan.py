import os

fox_root = r"G:\내 드라이브\[FoxConnect]"

for root, dirs, files in os.walk(fox_root):
    for d in dirs:
        if "01_가산" in d or "가산_대륭" in d:
            print("Found Gasan parent folder:", root)
            print("Found Gasan folder:", d)
