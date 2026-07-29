import subprocess
import sys

sys.stdout.reconfigure(encoding='utf-8')

git_exe = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
cwd = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

remote_url = "https://github.com/barista9980-cloud/AX-Work.git"

def run_git(args):
    cmd = [git_exe] + args
    print("Running:", " ".join(cmd))
    res = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    print("STDOUT:\n", res.stdout)
    if res.stderr:
        print("STDERR:\n", res.stderr)
    return res

print("Setting Remote Origin to GitHub...")
run_git(["remote", "remove", "origin"])
run_git(["remote", "add", "origin", remote_url])

print("\nChecking Remote Origin...")
run_git(["remote", "-v"])

print("\nPushing to GitHub main/master branch...")
res1 = run_git(["push", "-u", "origin", "master"])
if res1.returncode != 0:
    run_git(["push", "-u", "origin", "main", "--force"])
