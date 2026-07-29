import json
import re

with open(r"C:\Users\User\OneDrive\바탕 화면\업무_AX\dogok_extracted_text.json", "r", encoding="utf-8") as f:
    data = json.load(f)

parsed_contracts = []

for item in data:
    fname = item["filename"]
    text = item["full_text"]
    
    # Extract filename components
    # 강남_도곡로1길23_1층_01_전대차_[㈜폭스에듀-㈜에스앤에이치트레이딩]_(241101).pdf
    pat = r"^강남_도곡로1길23_([^_]+)_(\d{2})_([^_]+)_\[(.*?)\]_\((\d{6})\)\.pdf$"
    m = re.match(pat, fname)
    
    unit = ""
    seq = ""
    c_type = ""
    parties = ""
    c_date_raw = ""
    
    if m:
        unit, seq, c_type, parties, c_date_raw = m.groups()
    
    # Financial extraction
    # Look for 보증금, 월세/차임, 계약기간
    deposit_matches = re.findall(r'보\s*증\s*금[^\d]*?([\d\,]+(?:\s*만\s*원|\s*원)?)', text)
    rent_matches = re.findall(r'(?:월\s*임\s*대\s*료|차\s*임|월\s*세)[^\d]*?([\d\,]+(?:\s*만\s*원|\s*원)?)', text)
    period_matches = re.findall(r'(20\d{2}[\.\-\s년]\s*\d{1,2}[\.\-\s월]\s*\d{1,2}[일]?\s*부터\s*20\d{2}[\.\-\s년]\s*\d{1,2}[\.\-\s월]\s*\d{1,2}[일]?)', text)
    
    # Clean text to display snippet
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    parsed_contracts.append({
        "filename": fname,
        "unit": unit,
        "seq": seq,
        "contract_type": c_type,
        "parties": parties,
        "contract_date_raw": c_date_raw,
        "extracted_deposit": deposit_matches[:3],
        "extracted_rent": rent_matches[:3],
        "extracted_period": period_matches[:2],
        "snippet": lines[:15]
    })

with open(r"C:\Users\User\OneDrive\바탕 화면\업무_AX\dogok_summary.json", "w", encoding="utf-8") as f:
    json.dump(parsed_contracts, f, ensure_ascii=False, indent=2)

print("Parsed details for Dogok-ro contracts.")
