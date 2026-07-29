import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

GIT_EXE = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Using Git Executable:", GIT_EXE)

commit_msg = "feat: Finalize Enterprise Asset Framework (Real Estate, Vehicles & Insurance) with External Audit/IPO Compliant Master Excel Registers (.xlsx) and Standardized Upload Queue"

# 1. git add .
print("\n--- Running Git Add ---")
res_add = subprocess.run([GIT_EXE, "add", "."], cwd=REPO_DIR, capture_output=True, text=True)
print(res_add.stdout, res_add.stderr)

# 2. git status
print("\n--- Running Git Status ---")
res_status = subprocess.run([GIT_EXE, "status"], cwd=REPO_DIR, capture_output=True, text=True)
print(res_status.stdout)

# 3. git commit
print("\n--- Running Git Commit ---")
res_commit = subprocess.run([GIT_EXE, "commit", "-m", commit_msg], cwd=REPO_DIR, capture_output=True, text=True)
print(res_commit.stdout, res_commit.stderr)

# 4. git push
print("\n--- Running Git Push ---")
res_push = subprocess.run([GIT_EXE, "push", "origin", "main"], cwd=REPO_DIR, capture_output=True, text=True)
print(res_push.stdout, res_push.stderr)

if res_push.returncode == 0 or "Everything up-to-date" in res_push.stderr or "Everything up-to-date" in res_push.stdout:
    print("\n==========================================")
    print("SUCCESSFULLY COMMITTED AND PUSHED TO GITHUB!")
    print("==========================================")
else:
    print("\nRetrying Git Push without origin/main args...")
    res_push2 = subprocess.run([GIT_EXE, "push"], cwd=REPO_DIR, capture_output=True, text=True)
    print(res_push2.stdout, res_push2.stderr)
