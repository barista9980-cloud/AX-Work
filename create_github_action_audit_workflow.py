import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
workflows_dir = os.path.join(REPO_DIR, r".github\workflows")
os.makedirs(workflows_dir, exist_ok=True)

workflow_p = os.path.join(workflows_dir, "repo_audit_ci.yml")

workflow_content = """name: AX-Work GitHub Repository Audit & Symmetry CI

on:
  push:
    branches: [ "main" ]
  pull_request:
    branches: [ "main" ]

jobs:
  audit:
    runs-on: ubuntu-latest

    steps:
    - name: Checkout Code
      uses: actions/checkout@v3

    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.11'

    - name: Check Naming Symmetry & Secret Leakage
      run: |
        echo "=== 1. Checking Secret Leakage ==="
        ! grep -r "AIzaSy" . --exclude-dir=.git || exit 1
        
        echo "=== 2. Checking Corporate Data Leakage (.csv, .xlsx, .docx) ==="
        ! git ls-files | grep -E '\\.(csv|xlsx|docx|pdf)$' || exit 1
        
        echo "=== 3. Checking Template Naming Symmetry ==="
        test -f templates/Corporate_Real_Estate_Contract_Note_Template.md
        test -f templates/Corporate_Vehicle_Contract_Note_Template.md
        test -f templates/Corporate_Insurance_Contract_Note_Template.md
        
        echo "=== REPOSITORY AUDIT PASSED 100% ==="
"""

with open(workflow_p, "w", encoding="utf-8") as f:
    f.write(workflow_content)

print("Created GitHub Action Audit Workflow:", workflow_p)
