import json
import csv
import os

def export_to_csv():
    json_path = os.path.join(os.path.dirname(__file__), "real_estate_parsed_catalog.json")
    csv_path = os.path.join(os.path.dirname(__file__), "FoxConnect_부동산_자산대장_1차목록.csv")
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    fieldnames = [
        "연번", "권역", "물건지 및 호수", "차수", "계약유형", "당사 구분",
        "임대인(전대인) [원문]", "임대인 (통합)",
        "임차인(전차인) [원문]", "임차인 (통합)",
        "계약체결일", "파일명", "파일크기(KB)"
    ]
    
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(fieldnames)
        
        for idx, item in enumerate(data, 1):
            writer.writerow([
                idx,
                item.get("region", ""),
                item.get("property_info", ""),
                item.get("sequence", ""),
                item.get("contract_type", ""),
                item.get("our_role", ""),
                item.get("party_a_lessor", ""),
                item.get("party_a_normalized", ""),
                item.get("party_b_lessee", ""),
                item.get("party_b_normalized", ""),
                item.get("contract_date", ""),
                item.get("filename", ""),
                item.get("file_size_kb", "")
            ])
            
    print(f"Successfully exported {len(data)} items to {csv_path}")

if __name__ == "__main__":
    export_to_csv()
