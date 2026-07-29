import os

print("Searching for .git directories or git executables...")

for drive in ["C:\\Program Files", "C:\\Program Files (x86)", "C:\\Users\\User\\AppData\\Local"]:
    if os.path.exists(drive):
        for root, dirs, files in os.walk(drive):
            if "git.exe" in files:
                print("Found git.exe at:", os.path.join(root, "git.exe"))
                break

ax_dir = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
print("Contents of 업무_AX directory:")
for item in os.listdir(ax_dir):
    print(" -", item)
