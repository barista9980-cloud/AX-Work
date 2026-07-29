import os
import json
import csv
import sqlite3

BASE_DIR = os.path.dirname(__file__)
DB_PATH = os.path.join(BASE_DIR, "data", "real_estate_assets.db")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
GDRIVE_OUTPUT_DIR = r"G:\내 드라이브\[부동산자산] FoxConnect 계약 관리\04_생성_보고서"

os.makedirs(OUTPUT_DIR, exist_ok=True)
if os.path.exists(r"G:\내 드라이브\[부동산자산] FoxConnect 계약 관리"):
    os.makedirs(GDRIVE_OUTPUT_DIR, exist_ok=True)

def build_2025_snapshot():
    target_date = "2025-12-31"
    
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM real_estate_contracts
            WHERE contract_date <= ?
            ORDER BY contract_date ASC, region ASC, property_info ASC
        """, (target_date,))
        rows = [dict(r) for r in cursor.fetchall()]

    print(f"Total active/valid contracts as of {target_date}: {len(rows)}")

    # Save to local output
    filename = "FoxConnect_부동산_자산현황_2025년12월31일기준.csv"
    local_path = os.path.join(OUTPUT_DIR, filename)
    gdrive_path = os.path.join(GDRIVE_OUTPUT_DIR, filename)

    headers = [
        "연번", "기준일자", "계약일자", "권역", "물건지 및 호수", "차수",
        "계약유형", "당사 구분", "임대인(통합)", "임차인(통합)", "파일명"
    ]

    for p in [local_path, gdrive_path]:
        try:
            with open(p, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for idx, r in enumerate(rows, 1):
                    writer.writerow([
                        idx,
                        target_date,
                        r["contract_date"],
                        r["region"],
                        r["property_info"],
                        r["sequence"],
                        r["contract_type"],
                        r["our_role"],
                        r["party_a_normalized"],
                        r["party_b_normalized"],
                        r["filename"]
                    ])
            print(f"Report saved to: {p}")
        except Exception as e:
            print(f"Error saving to {p}: {e}")

    # Generate JSON summary
    summary_path = os.path.join(OUTPUT_DIR, "snapshot_20251231_summary.json")
    summary_data = {
        "as_of_date": target_date,
        "total_contracts": len(rows),
        "by_role": {},
        "by_region": {},
        "rows": rows
    }
    for r in rows:
        role = r["our_role"]
        reg = r["region"]
        summary_data["by_role"][role] = summary_data["by_role"].get(role, 0) + 1
        summary_data["by_region"][reg] = summary_data["by_region"].get(reg, 0) + 1

    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary_data, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    build_2025_snapshot()
