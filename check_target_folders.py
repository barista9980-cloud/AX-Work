import os

root_base = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차_전대차계약"

if not os.path.exists(root_base):
    # Try searching under FoxConnect
    fox_root = r"G:\내 드라이브\[FoxConnect]"
    for r, d, f in os.walk(fox_root):
        if "01_임대차" in r or "01_부동산_자산관리" in r:
            root_base = r
            break

print(f"Base path: {root_base}")
if os.path.exists(root_base):
    for item in os.listdir(root_base):
        print("  -", item)
