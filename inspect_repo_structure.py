import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Inspecting Repository Root Files in:", REPO_DIR)

py_files = []
md_files = []
other_files = []
dirs = []

for item in os.listdir(REPO_DIR):
    item_p = os.path.join(REPO_DIR, item)
    if os.path.isdir(item_p):
        if not item.startswith("."):
            dirs.append(item)
    else:
        if item.endswith(".py"):
            py_files.append(item)
        elif item.endswith(".md"):
            md_files.append(item)
        else:
            other_files.append(item)

print(f"\n[Directories] ({len(dirs)}):", dirs)
print(f"\n[Python Scripts] ({len(py_files)}):", sorted(py_files))
print(f"\n[Markdown Files] ({len(md_files)}):", sorted(md_files))
print(f"\n[Other Files] ({len(other_files)}):", sorted(other_files))
