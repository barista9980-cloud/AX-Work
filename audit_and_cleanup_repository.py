import os
import sys
import subprocess
import shutil

sys.stdout.reconfigure(encoding='utf-8')

GIT_EXE = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Executing Audit Deletion of Obsolete & Old Skill Files...")

files_to_delete = [
    "PREVIOUS_README.md",
    "test_border.docx",
    ".gemini/skills/real-estate-asset-manager/SKILL.md"
]

for rel_p in files_to_delete:
    abs_p = os.path.join(REPO_DIR, rel_p)
    print(f"\n--- Removing Obsolete File: {rel_p} ---")
    res_rm = subprocess.run([GIT_EXE, "rm", "-f", rel_p], cwd=REPO_DIR, capture_output=True, text=True)
    print(res_rm.stdout, res_rm.stderr)
    
    if os.path.exists(abs_p):
        try:
            os.remove(abs_p)
            print(f"  Deleted local file: {rel_p}")
        except Exception as e:
            print("  Note on local remove:", e)

# Remove empty .gemini directory if exists
gemini_dir = os.path.join(REPO_DIR, ".gemini")
if os.path.exists(gemini_dir):
    shutil.rmtree(gemini_dir, ignore_errors=True)
    print("  Removed obsolete .gemini skill directory.")

# Git status, commit, and push
commit_msg = "refactor: Audit repository and remove obsolete test docs, binary leftovers, and legacy skill files"

print("\n--- Running Git Add ---")
subprocess.run([GIT_EXE, "add", "-u"], cwd=REPO_DIR)

print("\n--- Running Git Commit ---")
res_commit = subprocess.run([GIT_EXE, "commit", "-m", commit_msg], cwd=REPO_DIR, capture_output=True, text=True)
print(res_commit.stdout, res_commit.stderr)

print("\n--- Running Git Push ---")
res_push = subprocess.run([GIT_EXE, "push", "origin", "main"], cwd=REPO_DIR, capture_output=True, text=True)
print(res_push.stdout, res_push.stderr)

if res_push.returncode == 0 or "Everything up-to-date" in res_push.stderr or "Everything up-to-date" in res_push.stdout:
    print("\n==========================================")
    print("REPOSITORY AUDIT & CLEANUP COMPLETED SUCCESSFULLY!")
    print("==========================================")
