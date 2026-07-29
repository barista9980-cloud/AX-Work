import os
import json
import sqlite3
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "real_estate_assets.db")
JSON_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "real_estate_parsed_catalog.json")

class RealEstateDB:
    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self.init_db()

    def init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS real_estate_contracts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    filename TEXT UNIQUE,
                    region TEXT,
                    property_info TEXT,
                    sequence TEXT,
                    contract_type TEXT,
                    our_role TEXT,
                    party_a_lessor TEXT,
                    party_a_normalized TEXT,
                    party_b_lessee TEXT,
                    party_b_normalized TEXT,
                    contract_date TEXT,
                    start_date TEXT,
                    end_date TEXT,
                    deposit INTEGER DEFAULT 0,
                    monthly_rent INTEGER DEFAULT 0,
                    file_size_kb REAL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()

    def sync_from_catalog(self, json_file=JSON_PATH):
        if not os.path.exists(json_file):
            print(f"JSON catalog not found: {json_file}")
            return

        with open(json_file, 'r', encoding='utf-8') as f:
            catalog = json.load(f)

        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            for item in catalog:
                cursor.execute("""
                    INSERT OR REPLACE INTO real_estate_contracts (
                        filename, region, property_info, sequence, contract_type,
                        our_role, party_a_lessor, party_a_normalized,
                        party_b_lessee, party_b_normalized, contract_date, file_size_kb
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    item.get("filename"),
                    item.get("region"),
                    item.get("property_info"),
                    item.get("sequence"),
                    item.get("contract_type"),
                    item.get("our_role"),
                    item.get("party_a_lessor"),
                    item.get("party_a_normalized"),
                    item.get("party_b_lessee"),
                    item.get("party_b_normalized"),
                    item.get("contract_date"),
                    item.get("file_size_kb", 0.0)
                ))
            conn.commit()
        print(f"Synced {len(catalog)} contracts into SQLite DB.")

    def get_snapshot_as_of_date(self, target_date_str):
        """
        Returns active contracts as of target_date_str (Format: 'YYYY-MM-DD').
        Contract is active if contract_date <= target_date_str
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM real_estate_contracts 
                WHERE contract_date <= ?
                ORDER BY region, property_info, sequence
            """, (target_date_str,))
            rows = cursor.fetchall()
            return [dict(row) for row in rows]

if __name__ == "__main__":
    db = RealEstateDB()
    db.sync_from_catalog()
    snapshot = db.get_snapshot_as_of_date("2024-12-31")
    print(f"Active contracts as of 2024-12-31: {len(snapshot)} items.")
