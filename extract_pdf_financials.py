import os
import re
import json
import glob
from pypdf import PdfReader

DOWNLOAD_DIR = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"

def extract_financials(filepath):
    try:
        reader = PdfReader(filepath)
        full_text = ""
        for page in reader.pages[:3]: # First 3 pages usually contain terms
            text = page.extract_text()
            if text:
                full_text += text + "\n"
                
        # Regex patterns for Korean real estate terms
        deposit_match = re.search(r'(보증금|전세금)[\s\:\;\=\-\|]*([일이삼사오육칠팔구십백천만억조\d\,\.\s]+원)', full_text)
        rent_match = re.search(r'(월세|차임|월임대료)[\s\:\;\=\-\|]*([일이삼사오육칠팔구십백천만억조\d\,\.\s]+원)', full_text)
        term_match = re.search(r'(임대기간|계약기간)[\s\:\;\=\-\|]*(\d{4}[\.\-\s년]+\d{1,2}[\.\-\s월]+\d{1,2}[\.\-\s일]*\s*~\s*\d{4}[\.\-\s년]+\d{1,2}[\.\-\s월]+\d{1,2}[\.\-\s일]*)', full_text)
        
        return {
            "has_text": len(full_text.strip()) > 50,
            "deposit_found": deposit_match.group(2).strip() if deposit_match else None,
            "rent_found": rent_match.group(2).strip() if rent_match else None,
            "term_found": term_match.group(2).strip() if term_match else None,
            "text_sample": full_text[:200].replace("\n", " ") if full_text else ""
        }
    except Exception as e:
        return {"error": str(e), "has_text": False}

def run_extraction():
    catalog_path = os.path.join(os.path.dirname(__file__), "real_estate_parsed_catalog.json")
    with open(catalog_path, "r", encoding="utf-8") as f:
        catalog = json.load(f)
        
    for item in catalog:
        filepath = os.path.join(DOWNLOAD_DIR, item["filename"])
        fin_data = extract_financials(filepath)
        item["extracted_text_details"] = fin_data
        
    output_path = os.path.join(os.path.dirname(__file__), "real_estate_full_analysis.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)
        
    print(f"Extraction finished. Output written to {output_path}")

if __name__ == "__main__":
    run_extraction()
