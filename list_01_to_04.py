import os

base_path = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차"

print("Listing items under 01_임대차:")
for item in os.listdir(base_path):
    item_p = os.path.join(base_path, item)
    if os.path.isdir(item_p):
        print(f" [DIR] {item}")
        for sub in os.listdir(item_p):
            print(f"       └── {sub}")
