import os
import re
import json
import glob

DOWNLOAD_DIR = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"

def normalize_company_name(name):
    clean_name = name.strip()
    if any(alias in clean_name for alias in ["폭스에듀", "폭스커넥트"]):
        return "(주)폭스커넥트 [구 (주)폭스에듀]"
    return clean_name

def determine_contract_role(party_a, party_b):
    norm_a = normalize_company_name(party_a)
    norm_b = normalize_company_name(party_b)
    
    if "(주)폭스커넥트" in norm_b and "(주)폭스커넥트" not in norm_a:
        return "당사 임차 (임차인/전차인)"
    elif "(주)폭스커넥트" in norm_a and "(주)폭스커넥트" not in norm_b:
        return "당사 전대/임대 (임대인/전대인)"
    elif "(주)폭스커넥트" in norm_a and "(주)폭스커넥트" in norm_b:
        return "당사 내부 계약"
    else:
        return "기타 계약"

def parse_filename(filename):
    name_without_ext = os.path.splitext(filename)[0]
    
    flexible_pattern = r"^([^_]+)_(.+?)_(\d{2})_([^_]+)_\[(.*?)\]_\((\d{6})\)"
    match_flex = re.match(flexible_pattern, name_without_ext)
    
    if match_flex:
        region, property_info, seq, contract_type, parties, raw_date = match_flex.groups()
        party_list = parties.split('-') if '-' in parties else [parties]
        party_a = party_list[0] if len(party_list) > 0 else ""
        party_b = party_list[1] if len(party_list) > 1 else ""
        
        party_a_norm = normalize_company_name(party_a)
        party_b_norm = normalize_company_name(party_b)
        role = determine_contract_role(party_a, party_b)
        
        return {
            "filename": filename,
            "region": region,
            "property_info": property_info,
            "sequence": seq,
            "contract_type": contract_type,
            "party_a_lessor": party_a,
            "party_a_normalized": party_a_norm,
            "party_b_lessee": party_b,
            "party_b_normalized": party_b_norm,
            "our_role": role,
            "contract_date": f"20{raw_date[:2]}-{raw_date[2:4]}-{raw_date[4:6]}",
            "status": "Success"
        }
        
    return {
        "filename": filename,
        "status": "Unparsed",
        "raw_name": name_without_ext
    }

def analyze_all_files():
    if not os.path.exists(DOWNLOAD_DIR):
        print(f"Directory not found: {DOWNLOAD_DIR}")
        return
        
    pdf_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.pdf"))
    results = []
    
    for filepath in pdf_files:
        filename = os.path.basename(filepath)
        data = parse_filename(filename)
        data["file_size_kb"] = round(os.path.getsize(filepath) / 1024, 1)
        results.append(data)
        
    output_path = os.path.join(os.path.dirname(__file__), "real_estate_parsed_catalog.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
        
    print(f"Total PDF Files Analyzed: {len(results)}")

if __name__ == "__main__":
    analyze_all_files()
