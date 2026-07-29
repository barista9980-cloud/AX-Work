import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Executing Secret Sanitization & Final End-of-Day Git Push...")

# 1. Sanitize Python scripts for any plain-text API keys or tokens
for root, dirs, files in os.walk(REPO_DIR):
    if ".git" in root:
        continue
    for f in files:
        if f.endswith(".py") or f.endswith(".md") or f.endswith(".json"):
            fp = os.path.join(root, f)
            try:
                with open(fp, "r", encoding="utf-8") as file:
                    content = file.read()
                
                # Check for sensitive patterns and sanitize
                sanitized = False
                if "SANITIZED_KEY" in content:
                    print(f"  [SANITIZING SECRET IN] {f}")
                    # Replace API keys if any
                    content = content.replace("SANITIZED_KEY", "SANITIZED_KEY")
                    sanitized = True
                
                if sanitized:
                    with open(fp, "w", encoding="utf-8") as file:
                        file.write(content)
            except Exception:
                pass

# 2. Git Status, Add, Commit, Push
commit_msg = "feat: Finalize Enterprise Asset Management Framework (Real Estate, Corporate Vehicles & Insurance) with External Audit/IPO Compliant Master Excel Registers (.xlsx) and Standardized Directory Structures"

try:
    print("\n--- Running Git Add ---")
    res_add = subprocess.run(["git", "add", "."], cwd=REPO_DIR, capture_output=True, text=True)
    print(res_add.stdout, res_add.stderr)

    print("\n--- Running Git Commit ---")
    res_commit = subprocess.run(["git", "commit", "-m", commit_msg], cwd=REPO_DIR, capture_output=True, text=True)
    print(res_commit.stdout, res_commit.stderr)

    print("\n--- Running Git Push ---")
    res_push = subprocess.run(["git", "push", "origin", "main"], cwd=REPO_DIR, capture_output=True, text=True)
    print(res_push.stdout, res_push.stderr)

    if res_push.returncode == 0 or "Everything up-to-date" in res_push.stderr or "Everything up-to-date" in res_push.stdout:
        print("\n==========================================")
        print("SUCCESSFULLY COMMITTED AND PUSHED TO GITHUB!")
        print("==========================================")
    else:
        # Retry push
        print("\nRetrying Git Push...")
        res_push2 = subprocess.run(["git", "push"], cwd=REPO_DIR, capture_output=True, text=True)
        print(res_push2.stdout, res_push2.stderr)

except Exception as e:
    print("Git Execution Error:", e)
