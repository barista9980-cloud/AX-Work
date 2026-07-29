import os
import re
import json
import csv
from pypdf import PdfReader

DOWNLOAD_DIR = r"C:\Users\User\Downloads\drive-download-20260723T054808Z-1-001"
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
GDRIVE_OUTPUT_DIR = r"G:\내 드라이브\[부동산자산] FoxConnect 계약 관리\04_생성_보고서"

os.makedirs(OUTPUT_DIR, exist_ok=True)
if os.path.exists(r"G:\내 드라이브\[부동산자산] FoxConnect 계약 관리"):
    os.makedirs(GDRIVE_OUTPUT_DIR, exist_ok=True)

def parse_start_date_and_terms(filepath, raw_date_str):
    """
    Attempts to extract start date, end date, deposit, rent from PDF text,
    falling back to filename date if not explicitly found in text.
    """
    start_date = f"20{raw_date_str[:2]}-{raw_date_str[2:4]}-{raw_date_str[4:6]}"
    end_date = "미정/확인필요"
    deposit = "-"
    rent = "-"
    
    try:
        reader = PdfReader(filepath)
        full_text = ""
        for page in reader.pages[:3]:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
                
        # Match pattern: YYYY년 MM월 DD일 ~ YYYY년 MM월 DD일
        term_match = re.search(r'(\d{4})[\.\s년\-]+(\d{1,2})[\.\s월\-]+(\d{1,2})[\.\s일]*\s*[~부터\s]+(\d{4})[\.\s년\-]+(\d{1,2})[\.\s월\-]+(\d{1,2})', full_text)
        if term_match:
            y1, m1, d1, y2, m2, d2 = term_match.groups()
            start_date = f"{y1}-{int(m1):02d}-{int(d1):02d}"
            end_date = f"{y2}-{int(m2):02d}-{int(d2):02d}"
            
        # Match deposit
        dep_match = re.search(r'(보증금|전세금)[\s\:\;\=\-\|]*([일이삼사오육칠팔구십백천만억조\d\,\.\s]+원)', full_text)
        if dep_match:
            deposit = dep_match.group(2).strip().replace("\n", "")
            
        # Match rent
        rent_m = re.search(r'(월세|차임|월임대료)[\s\:\;\=\-\|]*([일이삼사오육칠팔구십백천만억조\d\,\.\s]+원)', full_text)
        if rent_m:
            rent = rent_m.group(2).strip().replace("\n", "")
            
    except Exception as e:
        pass
        
    return start_date, end_date, deposit, rent

def generate_start_date_report():
    catalog_path = os.path.join(os.path.dirname(__file__), "real_estate_parsed_catalog.json")
    with open(catalog_path, "r", encoding="utf-8") as f:
        catalog = json.load(f)
        
    report_items = []
    
    for item in catalog:
        filename = item["filename"]
        filepath = os.path.join(DOWNLOAD_DIR, filename)
        
        # Extract filename raw date e.g. (260429) -> 260429
        raw_date_match = re.search(r'\((\d{6})\)', filename)
        raw_date = raw_date_match.group(1) if raw_date_match else "240101"
        
        s_date, e_date, deposit, rent = parse_start_date_and_terms(filepath, raw_date)
        
        # Extract year and month for grouping
        year_str = s_date.split("-")[0]
        month_str = s_date.split("-")[1] if len(s_date.split("-")) > 1 else "01"
        
        report_items.append({
            "start_year": year_str,
            "start_month": month_str,
            "start_date": s_date,
            "end_date": e_date,
            "region": item.get("region", ""),
            "property_info": item.get("property_info", ""),
            "sequence": item.get("sequence", ""),
            "contract_type": item.get("contract_type", ""),
            "our_role": item.get("our_role", ""),
            "lessor": item.get("party_a_normalized", ""),
            "lessee": item.get("party_b_normalized", ""),
            "deposit": deposit,
            "monthly_rent": rent,
            "filename": filename
        })
        
    # Sort chronologically by start_date
    report_items.sort(key=lambda x: (x["start_date"], x["region"], x["property_info"]))
    
    # Save CSV Report
    headers = [
        "연번", "계약시작일", "계약종료일", "시작연도", "시작월", "권역",
        "물건지 및 호수", "차수", "계약유형", "당사 구분", 
        "임대인(전대인)", "임차인(전차인)", "보증금(추출)", "월세(추출)", "파일명"
    ]
    
    csv_filename = "FoxConnect_부동산_계약시작일기준_현황보고서.csv"
    local_csv_path = os.path.join(OUTPUT_DIR, csv_filename)
    
    with open(local_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for idx, item in enumerate(report_items, 1):
            writer.writerow([
                idx,
                item["start_date"],
                item["end_date"],
                item["start_year"],
                item["start_month"],
                item["region"],
                item["property_info"],
                item["sequence"],
                item["contract_type"],
                item["our_role"],
                item["lessor"],
                item["lessee"],
                item["deposit"],
                item["monthly_rent"],
                item["filename"]
            ])
            
    print(f"Generated Start-Date Based Report: {local_csv_path} ({len(report_items)} items sorted chronologically)")
    
    # Sync to Google Drive output folder if exists
    if os.path.exists(GDRIVE_OUTPUT_DIR):
        gdrive_csv_path = os.path.join(GDRIVE_OUTPUT_DIR, csv_filename)
        with open(gdrive_csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for idx, item in enumerate(report_items, 1):
                writer.writerow([
                    idx,
                    item["start_date"],
                    item["end_date"],
                    item["start_year"],
                    item["start_month"],
                    item["region"],
                    item["property_info"],
                    item["sequence"],
                    item["contract_type"],
                    item["our_role"],
                    item["lessor"],
                    item["lessee"],
                    item["deposit"],
                    item["monthly_rent"],
                    item["filename"]
                ])
        print(f"Synced report to Google Drive: {gdrive_csv_path}")

if __name__ == "__main__":
    generate_start_date_report()
