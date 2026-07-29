import os

parent_dir = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차"

# Get full matching directory path
fox_root = r"G:\내 드라이브\[FoxConnect]"
target_parent = None

for root, dirs, files in os.walk(fox_root):
    for d in dirs:
        if "01_가산" in d or "대륭포스트타워" in d:
            target_parent = root
            break
    if target_parent:
        break

print(f"Target Parent Directory: {target_parent}")

folders_01_to_04 = []

if target_parent and os.path.exists(target_parent):
    all_sub = os.listdir(target_parent)
    all_sub.sort()
    for sub in all_sub:
        full_p = os.path.join(target_parent, sub)
        if os.path.isdir(full_p):
            print(f"\n[DIR] {sub}")
            units = os.listdir(full_p)
            units.sort()
            for u in units:
                u_p = os.path.join(full_p, u)
                if os.path.isdir(u_p):
                    print(f"   └── [UNIT DIR] {u}")
