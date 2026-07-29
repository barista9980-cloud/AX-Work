import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"

print("Auditing all tracked files for internal reference consistency & naming symmetry...")

old_str = "Corporate_Real_Estate_Contract_Note_Template.md"
new_str = "Corporate_Corporate_Real_Estate_Contract_Note_Template.md"

replaced_files = []

for root, dirs, files in os.walk(REPO_DIR):
    if ".git" in root:
        continue
    for f in files:
        if f.endswith(".md") or f.endswith(".py") or f.endswith(".json"):
            fp = os.path.join(root, f)
            try:
                with open(fp, "r", encoding="utf-8") as file:
                    content = file.read()
                
                if old_str in content:
                    print(f"  [UPDATING REFERENCE IN] {f}")
                    new_content = content.replace(old_str, new_str)
                    with open(fp, "w", encoding="utf-8") as file:
                        file.write(new_content)
                    replaced_files.append(f)
            except Exception as e:
                pass

print(f"\nAudit complete! Updated references in {len(replaced_files)} files.")
