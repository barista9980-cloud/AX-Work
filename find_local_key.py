import os
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

log_file = r"C:\Users\User\.gemini\antigravity-cli\brain\8aa87ffe-68a4-4102-991e-dc1cfb195426\.system_generated\logs\transcript_full.jsonl"

found_keys = []
if os.path.exists(log_file):
    with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            m = re.findall(r"SANITIZED_KEY[a-zA-Z0-9_\-]+", line)
            if m:
                found_keys.extend(m)

found_keys = list(set(found_keys))
print(f"Found {len(found_keys)} API Keys in conversation logs:")
for k in found_keys:
    print(" - Key:", k)

if found_keys:
    with open("local_api_key.txt", "w", encoding="utf-8") as fk:
        fk.write(found_keys[0])
