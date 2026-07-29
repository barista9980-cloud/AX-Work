import os
import sys
import docx
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
VEHICLE_1TO1_BASE = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\02_차량_자산관리\01_차량계약_리스_렌트")

print("Inspecting ALL 10 1:1 Vehicle Folders & Word Notes in:", VEHICLE_1TO1_BASE)

vehicle_dirs = sorted([d for d in os.listdir(VEHICLE_1TO1_BASE) if os.path.isdir(os.path.join(VEHICLE_1TO1_BASE, d))])
print(f"Found {len(vehicle_dirs)} 1:1 Vehicle Folders:", vehicle_dirs)

vehicle_master_records = []

for v_dir in vehicle_dirs:
    dir_p = os.path.join(VEHICLE_1TO1_BASE, v_dir)
    print(f"\n==========================================")
    print(f"Folder: {v_dir}")
    print(f"==========================================")
    
    files_in_dir = os.listdir(dir_p)
    docx_files = [f for f in files_in_dir if f.endswith(".docx")]
    pdf_files = [f for f in files_in_dir if f.endswith(".pdf")]
    
    note_info = {}
    
    if docx_files:
        doc_p = os.path.join(dir_p, docx_files[0])
        print("  Reading Word Note:", docx_files[0])
        doc = docx.Document(doc_p)
        for t in doc.tables:
            for r in t.rows:
                row_txt = [c.text.strip().replace("\n", " ") for c in r.cells]
                print("    TBL:", " | ".join(row_txt))
    else:
        print("  No docx note found!")
        
    print("  PDF Files:", pdf_files)
