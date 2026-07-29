import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

GIT_EXE = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
templates_dir = os.path.join(REPO_DIR, "templates")

old_name = "Real_Estate_Contract_Note_Template.md"
new_name = "Corporate_Real_Estate_Contract_Note_Template.md"

old_p = os.path.join(templates_dir, old_name)
new_p = os.path.join(templates_dir, new_name)

print("Standardizing Real Estate Template File Name for 100% Naming Symmetry...")

if os.path.exists(old_p):
    # Git mv old_name new_name
    res_mv = subprocess.run([GIT_EXE, "mv", f"templates/{old_name}", f"templates/{new_name}"], cwd=REPO_DIR, capture_output=True, text=True)
    print(res_mv.stdout, res_mv.stderr)
    
    if not os.path.exists(new_p) and os.path.exists(old_p):
        os.rename(old_p, new_p)

print("Template directory listing:")
for item in os.listdir(templates_dir):
    print("  -", item)

# Git commit and push
commit_msg = "refactor: Standardize Real Estate template filename to Corporate_Real_Estate_Contract_Note_Template.md for 100% naming symmetry"

print("\n--- Running Git Add ---")
subprocess.run([GIT_EXE, "add", "-A"], cwd=REPO_DIR)

print("\n--- Running Git Commit ---")
res_commit = subprocess.run([GIT_EXE, "commit", "-m", commit_msg], cwd=REPO_DIR, capture_output=True, text=True)
print(res_commit.stdout, res_commit.stderr)

print("\n--- Running Git Push ---")
res_push = subprocess.run([GIT_EXE, "push", "origin", "main"], cwd=REPO_DIR, capture_output=True, text=True)
print(res_push.stdout, res_push.stderr)

if res_push.returncode == 0 or "Everything up-to-date" in res_push.stderr or "Everything up-to-date" in res_push.stdout:
    print("\n==========================================")
    print("TEMPLATE SYMMETRY RENAMED & PUSHED SUCCESSFULLY!")
    print("==========================================")
