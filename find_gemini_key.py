import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("Checking environment variables for Gemini Key:")
for k, v in os.environ.items():
    if "API" in k or "GEMINI" in k or "GOOGLE" in k or "KEY" in k:
        print(f"  {k} = {v[:10]}...")

# Search in OneDrive script files for key variable
sc_dir = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
for f in os.listdir(sc_dir):
    if f.endswith(".py") and "key" in f.lower():
        print("Found key script:", f)
