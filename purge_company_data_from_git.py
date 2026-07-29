import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

GIT_EXE = r"C:\Users\User\AppData\Local\GitHubDesktop\app-3.6.3\resources\app\git\cmd\git.exe"
REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Executing Company Data Purge & Strict .gitignore Update...")

# 1. Update .gitignore to ignore all data deliverables (*.csv, *.xlsx, *.docx, *.pdf, *.json, *.txt, etc.)
gitignore_p = os.path.join(REPO_DIR, ".gitignore")
gitignore_content = """# Environments & Python
__pycache__/
*.py[cod]
*$py.class
.env
venv/
ENV/

# Data Deliverables & Corporate Content (NEVER UPLOAD TO GITHUB)
*.csv
*.xlsx
*.xls
*.docx
*.doc
*.pdf
*.png
*.jpg
*.jpeg
*.json
*.txt
!mcp.json.template

# Specific Corporate File Patterns
FoxConnect_*
[외감_IPO대비]*
부동산_*
차량_*
보험_*

# System & Local Configs
local_api_key.txt
.DS_Store
Thumbs.db
"""

with open(gitignore_p, "w", encoding="utf-8") as f:
    f.write(gitignore_content)
print("  Updated .gitignore with strict data exclusion rules!")

# 2. Git rm FoxConnect_부동산_자산대장_1차목록.csv
csv_name = "FoxConnect_부동산_자산대장_1차목록.csv"
csv_p = os.path.join(REPO_DIR, csv_name)

print(f"\n--- Removing {csv_name} from Git Repository ---")
res_rm = subprocess.run([GIT_EXE, "rm", "-f", csv_name], cwd=REPO_DIR, capture_output=True, text=True)
print(res_rm.stdout, res_rm.stderr)

if os.path.exists(csv_p):
    try:
        os.remove(csv_p)
        print(f"  Removed local copy of {csv_name} from repo root.")
    except Exception as e:
        print("  Note on local remove:", e)

# 3. Git add .gitignore & Commit
commit_msg = "security: Purge corporate CSV data deliverable and enforce strict .gitignore rules excluding all company data files"

print("\n--- Running Git Add ---")
subprocess.run([GIT_EXE, "add", ".gitignore"], cwd=REPO_DIR)
subprocess.run([GIT_EXE, "add", "-u"], cwd=REPO_DIR)

print("\n--- Running Git Commit ---")
res_commit = subprocess.run([GIT_EXE, "commit", "-m", commit_msg], cwd=REPO_DIR, capture_output=True, text=True)
print(res_commit.stdout, res_commit.stderr)

print("\n--- Running Git Push ---")
res_push = subprocess.run([GIT_EXE, "push", "origin", "main"], cwd=REPO_DIR, capture_output=True, text=True)
print(res_push.stdout, res_push.stderr)

if res_push.returncode == 0 or "Everything up-to-date" in res_push.stderr or "Everything up-to-date" in res_push.stdout:
    print("\n==========================================")
    print("CORPORATE DATA PURGED & GITHUB CLEANED SUCCESSFULLY!")
    print("==========================================")
