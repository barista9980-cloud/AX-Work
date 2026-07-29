import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

GIT_EXE = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Fetching COMPLETE list of files tracked by Git in AX-Work...")

res = subprocess.run([GIT_EXE, "ls-files"], cwd=REPO_DIR, capture_output=True, text=True)
tracked_files = [f.strip() for f in res.stdout.split("\n") if f.strip()]

print(f"\nTotal Tracked Files in Git ({len(tracked_files)} files):")
for idx, tf in enumerate(tracked_files, 1):
    file_p = os.path.join(REPO_DIR, tf)
    size = os.path.getsize(file_p) if os.path.exists(file_p) else 0
    print(f"[{idx:02d}] {tf:<60} ({size} bytes)")
