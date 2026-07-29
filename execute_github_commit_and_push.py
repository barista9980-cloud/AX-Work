import os
import subprocess
import sys

sys.stdout.reconfigure(encoding='utf-8')

git_exe = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
cwd = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Using Git Executable:", git_exe)

def run_git(args):
    cmd = [git_exe] + args
    print("Running:", " ".join(cmd))
    res = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    print("STDOUT:\n", res.stdout)
    if res.stderr:
        print("STDERR:\n", res.stderr)
    return res

# 1. Check if git repo exists
if not os.path.exists(os.path.join(cwd, ".git")):
    print("Initializing Git Repository...")
    run_git(["init"])

# 2. Check git status
run_git(["status"])

# 3. Git add all changes
print("\n--- ADDING ALL CHANGED FILES TO GIT ---")
run_git(["add", "."])

# 4. Commit changes
commit_msg = "docs: 부동산 계약관리노트 최상위 마스터 템플릿 개편 및 Vision LLM AX 파이프라인 구축"
print(f"\n--- COMMITTING CHANGES: {commit_msg} ---")
run_git(["commit", "-m", commit_msg])

# 5. Check git remote
print("\n--- CHECKING REMOTE ---")
res_remote = run_git(["remote", "-v"])

if "origin" in res_remote.stdout:
    print("\n--- PUSHING TO GITHUB ORIGIN ---")
    run_git(["push", "origin", "main"])
    run_git(["push", "origin", "master"])
else:
    print("\n[NOTE] Remote 'origin' is not set yet. Local Git commit completed successfully!")
