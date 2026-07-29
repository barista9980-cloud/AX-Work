import os

print("Searching for modified 402_403 docx file...")

search_locations = [
    r"G:\내 드라이브\[FoxConnect]",
    r"C:\Users\User\OneDrive\바탕 화면",
    r"C:\Users\User\Downloads"
]

found_files = []

for loc in search_locations:
    if os.path.exists(loc):
        for root, dirs, files in os.walk(loc):
            for f in files:
                if f.endswith(".docx") and "402" in f:
                    full_p = os.path.join(root, f)
                    mtime = os.path.getmtime(full_p)
                    found_files.append((mtime, full_p))

found_files.sort(reverse=True)

for mtime, path in found_files:
    import datetime
    dt = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{dt}] {path}")
