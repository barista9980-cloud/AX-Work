import os
import docx

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

print("=== CHECKING OVERNIGHT VISION LLM PROCESSING RESULTS ===")

found_notes = []

for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    for f in files:
        if f.endswith(".docx") and "계약관리노트" in f:
            full_path = os.path.join(root, f)
            folder_name = os.path.basename(root)
            mtime = os.path.getmtime(full_path)
            
            # Read first few table values
            try:
                doc = docx.Document(full_path)
                t0 = doc.tables[0]
                period_str = t0.rows[2].cells[1].text.strip()
                deposit_str = t0.rows[3].cells[1].text.strip()
                rent_str = t0.rows[3].cells[3].text.strip()
                
                found_notes.append({
                    "folder": folder_name,
                    "filename": f,
                    "period": period_str,
                    "deposit": deposit_str,
                    "rent": rent_str,
                    "mtime": mtime
                })
            except Exception as e:
                print(f"Error reading {f}: {e}")

print(f"\nTotal Contract Notes Found: {len(found_notes)}")
print("\nSample Processed Contract Notes:")
for item in found_notes[:10]:
    print(f" - [{item['folder']}] {item['filename']} | 임대기간: {item['period']} | 보증금: {item['deposit']} | 월세: {item['rent']}")
