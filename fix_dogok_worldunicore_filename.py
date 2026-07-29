import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

dogok_path = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\02_강남_도곡로1길23\지하1층,1층,2층,3층"

old_f = "강남_도곡로1길23_1층_03_전대차_[㈜폭스에듀-㈜월드유니코어]_(250821).pdf"
new_f = "07_전대차_강남_도곡로1길23_1층_[㈜폭스에듀-㈜월드유니코어]_(250821).pdf"

old_p = os.path.join(dogok_path, old_f)
new_p = os.path.join(dogok_path, new_f)

if os.path.exists(old_p):
    try:
        os.rename(old_p, new_p)
        print("Successfully renamed:", old_f, "->", new_f)
    except Exception as e:
        print("Error renaming:", e)
else:
    print("Old file not found:", old_p)

# Verify all files in dogok folder
print("\nUpdated Files in Dogok folder:")
files = sorted(os.listdir(dogok_path))
for f in files:
    print(" -", f)
