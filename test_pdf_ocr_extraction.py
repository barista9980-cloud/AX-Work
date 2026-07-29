import os

print("Checking available PDF extraction libraries in Python...")

pdf_tools = {}

try:
    import pypdf
    pdf_tools["pypdf"] = True
except ImportError:
    pdf_tools["pypdf"] = False

try:
    import pdfplumber
    pdf_tools["pdfplumber"] = True
except ImportError:
    pdf_tools["pdfplumber"] = False

try:
    import fitz # PyMuPDF
    pdf_tools["fitz"] = True
except ImportError:
    pdf_tools["fitz"] = False

try:
    from rapidocr_onnxruntime import RapidOCR
    pdf_tools["rapidocr"] = True
except ImportError:
    pdf_tools["rapidocr"] = False

try:
    import pdf2image
    pdf_tools["pdf2image"] = True
except ImportError:
    pdf_tools["pdf2image"] = False

print("Available PDF Tools:", pdf_tools)

# Find sample PDF in 402_403호
FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

pdf_path = None
for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    if "402_403" in root:
        for f in files:
            if f.endswith(".pdf"):
                pdf_path = os.path.join(root, f)
                break
    if pdf_path:
        break

print("Found PDF path:", pdf_path)
