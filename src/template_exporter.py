import os
import csv
import json
import sqlite3
from datetime import datetime

class TemplateExporter:
    def __init__(self, db_path=None):
        self.base_dir = os.path.dirname(os.path.dirname(__file__))
        self.templates_dir = os.path.join(self.base_dir, "templates")
        self.output_dir = os.path.join(self.base_dir, "output")
        self.db_path = db_path or os.path.join(self.base_dir, "data", "real_estate_assets.db")
        
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.templates_dir, exist_ok=True)

    def generate_snapshot_report(self, as_of_date="2024-12-31", output_name=None):
        """
        Generates a standardized report for a target snapshot date.
        """
        from db_manager import RealEstateDB
        db = RealEstateDB(self.db_path)
        records = db.get_snapshot_as_of_date(as_of_date)
        
        if not output_name:
            clean_date = as_of_date.replace("-", "")
            output_name = f"부동산_자산현황_스냅샷_{clean_date}.csv"
            
        out_filepath = os.path.join(self.output_dir, output_name)
        
        fieldnames = [
            "연번", "기준일자", "권역", "물건지 및 호수", "차수", "계약유형",
            "당사 구분", "임대인(통합)", "임차인(통합)", "계약체결일", "파일명"
        ]
        
        with open(out_filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(fieldnames)
            
            for idx, r in enumerate(records, 1):
                writer.writerow([
                    idx,
                    as_of_date,
                    r["region"],
                    r["property_info"],
                    r["sequence"],
                    r["contract_type"],
                    r["our_role"],
                    r["party_a_normalized"],
                    r["party_b_normalized"],
                    r["contract_date"],
                    r["filename"]
                ])
                
        print(f"Generated snapshot report ({as_of_date}): {out_filepath} ({len(records)} records)")
        return out_filepath

if __name__ == "__main__":
    exporter = TemplateExporter()
    exporter.generate_snapshot_report("2022-12-31")
    exporter.generate_snapshot_report("2024-12-31")
    exporter.generate_snapshot_report("2026-12-31")
