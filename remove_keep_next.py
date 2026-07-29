import os
import docx

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

target_parent = None
for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    for d in dirs:
        if "01_가산" in d or "대륭포스트타워" in d:
            target_parent = root
            break
    if target_parent:
        break

print(f"Target Parent: {target_parent}")

target_folders_prefix = ["01_가산", "02_강남", "03_광명", "04_광명"]
sub_dirs = os.listdir(target_parent)
sub_dirs.sort()

count = 0

for s_dir in sub_dirs:
    if not any(s_dir.startswith(pref) for pref in target_folders_prefix):
        continue
    folder_path = os.path.join(target_parent, s_dir)
    if not os.path.isdir(folder_path):
        continue

    for u_dir in os.listdir(folder_path):
        u_path = os.path.join(folder_path, u_dir)
        if not os.path.isdir(u_path):
            continue

        docx_files = [f for f in os.listdir(u_path) if f.endswith(".docx")]
        for df in docx_files:
            df_path = os.path.join(u_path, df)
            try:
                doc = docx.Document(df_path)
                changed = False
                for p in doc.paragraphs:
                    # Disable keep_with_next
                    if p.paragraph_format.keep_with_next:
                        p.paragraph_format.keep_with_next = False
                        changed = True
                    # Remove any keepNext elements explicitly from XML
                    pPr = p._p.get_or_add_pPr()
                    kn = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}keepNext')
                    if kn is not None:
                        pPr.remove(kn)
                        changed = True

                if changed:
                    doc.save(df_path)
                    print(f"Removed keepNext (left-margin black dot) from: {df}")
                    count += 1
            except Exception as e:
                print(f"Error processing {df}: {e}")

print(f"Removed keepNext indicator from {count} docx files.")
