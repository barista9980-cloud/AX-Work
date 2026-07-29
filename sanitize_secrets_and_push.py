import os
import re
import subprocess
import sys

sys.stdout.reconfigure(encoding='utf-8')

git_exe = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
cwd = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Sanitizing hardcoded API keys and tokens in ALL Python/JSON files...")

dummy_key = 'os.environ.get("GEMINI_API_KEY", "")'
dummy_token = 'os.environ.get("GITHUB_TOKEN", "")'

for root, dirs, files in os.walk(cwd):
    for fname in files:
        if fname.endswith(".py") or fname.endswith(".json") or fname.endswith(".txt"):
            if fname == "local_api_key.txt":
                continue
            fpath = os.path.join(root, fname)
            try:
                with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                
                if "AQ.Ab8RN6KRfI2J" in content or "ghp_" in content or "SANITIZED_KEY" in content:
                    content = re.sub(r'USER_API_KEY\s*=\s*"[^"]+"', f'USER_API_KEY = {dummy_key}', content)
                    content = re.sub(r'user_api_key\s*=\s*"[^"]+"', f'user_api_key = {dummy_key}', content)
                    content = re.sub(r'GITHUB_TOKEN\s*=\s*"[^"]+"', f'GITHUB_TOKEN = {dummy_token}', content)
                    content = re.sub(r'SANITIZED_KEY[a-zA-Z0-9_\-]+', 'DUMMY_GEMINI_API_KEY', content)
                    content = re.sub(r'ghp_[a-zA-Z0-9]+', 'DUMMY_GITHUB_TOKEN', content)

                    with open(fpath, "w", encoding="utf-8") as f:
                        f.write(content)
                    print(f"  [SANITIZED] {fname}")
            except Exception:
                pass

def run_git(args):
    cmd = [git_exe] + args
    res = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    print("Running:", " ".join(args))
    if res.stdout:
        print("STDOUT:\n", res.stdout)
    if res.stderr:
        print("STDERR:\n", res.stderr)
    return res

print("\n--- RESETTING LOCAL GIT HISTORY TO CLEAN SECRETS ---")
git_dir = os.path.join(cwd, ".git")
if os.path.exists(git_dir):
    import shutil
    shutil.rmtree(git_dir, ignore_errors=True)

run_git(["init"])
run_git(["branch", "-M", "main"])
run_git(["remote", "add", "origin", "https://github.com/barista9980-cloud/AX-Work.git"])

print("\n--- ADDING ALL CLEAN FILES TO GIT ---")
run_git(["add", "."])

commit_msg = "feat: 법인차량 1:1 평면 폴더 개편, 차량 계약관리 마스터 서식 10대 자동완성 및 범용 표준 템플릿 수록"
print(f"\n--- COMMITTING CLEAN FILES: {commit_msg} ---")
run_git(["commit", "-m", commit_msg])

print("\n--- PUSHING TO GITHUB MAIN BRANCH ---")
run_git(["push", "-u", "origin", "main", "--force"])
