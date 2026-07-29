import os
import subprocess

print("Finding git.exe on Windows...")

possible_paths = [
    r"C:\Program Files\Git\cmd\git.exe",
    r"C:\Program Files (x86)\Git\cmd\git.exe",
    r"C:\Users\User\AppData\Local\Programs\Git\cmd\git.exe",
    r"C:\Users\User\AppData\Local\Git\cmd\git.exe",
    r"C:\Program Files\Git\bin\git.exe"
]

git_bin = None
for p in possible_paths:
    if os.path.exists(p):
        git_bin = p
        break

if not git_bin:
    # search path
    try:
        res = subprocess.run(["where", "git"], capture_output=True, text=True)
        if res.returncode == 0:
            git_bin = res.stdout.splitlines()[0].strip()
    except Exception:
        pass

print("Git binary path found:", git_bin)

cwd = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

if git_bin:
    # Check if git repo
    is_repo = os.path.exists(os.path.join(cwd, ".git"))
    print("Is git repo in 업무_AX:", is_repo)
    
    if not is_repo:
        print("Initializing git repo in 업무_AX...")
        subprocess.run([git_bin, "init"], cwd=cwd)
        
    res = subprocess.run([git_bin, "status"], cwd=cwd, capture_output=True, text=True)
    print("Git status:\n", res.stdout)
else:
    print("Git binary not found on machine.")
