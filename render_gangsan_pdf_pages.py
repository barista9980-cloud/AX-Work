import os
import sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"
GANGSAN_DIR = os.path.join(FOXCONNECT_ROOT, r"[총무]업무\01_부동산_자산관리\01_임대차계약\22_강남_도곡로1길23_지하1층,1층,2층,3층")

pdf1 = os.path.join(GANGSAN_DIR, "01_최초임대차_강남_도곡로1길23_전층_[박재윤-㈜폭스에듀]_(241107).pdf")
pdf8 = os.path.join(GANGSAN_DIR, "08_변경계약_강남_도곡로1길23_전층_[유한회사 청송(박재윤)-㈜폭스에듀]_(250901).pdf")

output_dir = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\gangsan_pages"
os.makedirs(output_dir, exist_ok=True)

print("Rendering Gangsan Contract PDF pages to PNG...")

for pdf_path, prefix in [(pdf1, "pdf1"), (pdf8, "pdf8")]:
    if os.path.exists(pdf_path):
        doc = fitz.open(pdf_path)
        for i in range(min(5, len(doc))):
            page = doc[i]
            pix = page.get_pixmap(dpi=200)
            out_img = os.path.join(output_dir, f"{prefix}_page_{i+1}.png")
            pix.save(out_img)
            print(f"  Saved: {out_img}")

print("Rendering complete!")
