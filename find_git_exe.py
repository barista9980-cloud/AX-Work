import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

possible_paths = [
    r"C:\Program Files\Git\cmd\git.exe",
    r"C:\Program Files\Git\bin\git.exe",
    r"C:\Program Files (x86)\Git\cmd\git.exe",
    os.path.expanduser(r"~\AppData\Local\Programs\Git\cmd\git.exe"),
    r"C:\Users\User\AppData\Local\Programs\Git\cmd\git.exe",
    r"C:\Users\User\AppData\Local\GitHubDesktop\app-*\resources\app\git\cmd\git.exe"
]

print("Searching for git.exe...")

found_git = None
for p in possible_paths:
    if os.path.exists(p):
        found_git = p
        print("FOUND GIT AT:", p)
        break

if not found_git:
    # search C drive
    for root, dirs, files in os.walk(r"C:\Program Files"):
        if "git.exe" in files:
            found_git = os.path.join(root, "git.exe")
            print("FOUND GIT AT:", found_git)
            break

if not found_git:
    for root, dirs, files in os.walk(os.path.expanduser("~")):
        if "git.exe" in files:
            found_git = os.path.join(root, "git.exe")
            print("FOUND GIT AT:", found_git)
            break

print("Final Git Path:", found_git)
