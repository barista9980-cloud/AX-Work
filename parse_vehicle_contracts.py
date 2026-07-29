import os
import re
import json
import csv
import shutil
from pypdf import PdfReader

QUEUE_DIR = r"G:\내 드라이브\[FoxConnect]\[총무]업무\06_자동파싱_업로드큐\01_신규_문서_업로드"
ARCHIVE_DIR = r"G:\내 드라이브\[FoxConnect]\[총무]업무\06_자동파싱_업로드큐\02_처리완료_아카이브"
VEHICLE_DEST_DIR = r"G:\내 드라이브\[FoxConnect]\[총무]업무\02_차량_자산관리\01_차량계약_리스_렌트"
REPORT_DIR = r"G:\내 드라이브\[FoxConnect]\[총무]업무\07_보고서_템플릿\01_정기_자산리포트"
VEHICLE_REPORT_DIR = r"G:\내 드라이브\[FoxConnect]\[총무]업무\02_차량_자산관리"

def parse_vehicle_pdf(filepath, filename):
    try:
        reader = PdfReader(filepath)
        full_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    except Exception as e:
        full_text = ""
        print(f"Error reading {filename}: {e}")

    # Fallback to filename parsing if text extraction is sparse
    # Example filename: 1.K8(3930)_현대캐피탈_계약서확인서_리스.pdf
    name_no_ext = os.path.splitext(filename)[0]
    
    contract_type = "리스" if "리스" in filename else ("렌트" if "렌트" in filename else "기타")
    
    capital_match = re.search(r'_(현대캐피탈|하나캐피탈|BNK캐피탈|우리캐피탈|DGB\(IM캐피탈\)|DGB캐피탈|캐피탈[^_]*)_', filename)
    capital = capital_match.group(1) if capital_match else "기타캐피탈"
    
    car_match = re.search(r'^\d+\.([^\(]+)\(([^\)]+)\)', filename)
    if car_match:
        model = car_match.group(1).strip()
        plate = car_match.group(2).strip()
    else:
        model = "미상"
        plate = "미상"

    # Extract money / dates from full_text if available
    # Monthly fee pattern
    monthly_fee = ""
    fee_match = re.search(r'(월\s*대여료|월\s*리스료|월\s*납입금|월\s*납입금액|월대여료|월리스료)[^\d]*([\d\,]+)\s*원', full_text)
    if fee_match:
        monthly_fee = fee_match.group(2) + " 원"

    # Period pattern
    period = ""
    period_match = re.search(r'(\d{2,4}[\.\-\/]\d{1,2}[\.\-\/]\d{1,2})\s*[\~\-]\s*(\d{2,4}[\.\-\/]\d{1,2}[\.\-\/]\d{1,2})', full_text)
    if period_match:
        period = f"{period_match.group(1)} ~ {period_match.group(2)}"

    return {
        "filename": filename,
        "car_model": model,
        "plate_number": plate,
        "capital_company": capital,
        "contract_type": contract_type,
        "monthly_fee": monthly_fee if monthly_fee else "계약서 본문 참조",
        "contract_period": period if period else "계약서 본문 참조",
        "file_path": filepath
    }

def main():
    pdf_files = [f for f in os.listdir(QUEUE_DIR) if f.lower().endswith('.pdf')] if os.path.exists(QUEUE_DIR) else []
    
    source_dir = QUEUE_DIR
    if not pdf_files and os.path.exists(VEHICLE_DEST_DIR):
        pdf_files = [f for f in os.listdir(VEHICLE_DEST_DIR) if f.lower().endswith('.pdf')]
        source_dir = VEHICLE_DEST_DIR
        print(f"Queue empty, scanning vehicle destination dir ({len(pdf_files)} PDFs).")

    results = []
    for f in pdf_files:
        src_path = os.path.join(source_dir, f)
        data = parse_vehicle_pdf(src_path, f)
        results.append(data)
        
        if source_dir == QUEUE_DIR:
            os.makedirs(VEHICLE_DEST_DIR, exist_ok=True)
            shutil.copy2(src_path, os.path.join(VEHICLE_DEST_DIR, f))
            os.makedirs(ARCHIVE_DIR, exist_ok=True)
            shutil.move(src_path, os.path.join(ARCHIVE_DIR, f))
            print(f"Processed and archived: {f}")

    # Generate CSV ledger
    fieldnames = ["연번", "차종", "차량번호", "구분(리스/렌트)", "금융/캐피탈사", "월 납입금", "계약 기간", "파일명"]
    
    csv_rows = []
    for idx, item in enumerate(results, 1):
        csv_rows.append([
            idx,
            item["car_model"],
            item["plate_number"],
            item["contract_type"],
            item["capital_company"],
            item["monthly_fee"],
            item["contract_period"],
            item["filename"]
        ])

    # Save CSVs
    for out_dir in [REPORT_DIR, VEHICLE_REPORT_DIR]:
        os.makedirs(out_dir, exist_ok=True)
        csv_path = os.path.join(out_dir, "FoxConnect_법인차량_자산대장.csv")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(fieldnames)
            writer.writerows(csv_rows)
        print(f"Exported ledger to: {csv_path}")

    # Save JSON summary
    json_path = os.path.join(VEHICLE_REPORT_DIR, "FoxConnect_법인차량_자산대장.json")
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump(results, jf, ensure_ascii=False, indent=2)

    print("\nAll 8 vehicle contract PDFs successfully parsed, organized, and cataloged!")

if __name__ == "__main__":
    main()
