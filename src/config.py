"""
Universal Configuration for Corporate Asset Management Framework
Supports Environment Variables, config.json, and Interactive CLI Setup Wizard.
"""
import os
import sys
import json

CONFIG_JSON_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.json")

def load_user_config():
    """Load configuration from config.json or environment variables."""
    config_data = {}
    if os.path.exists(CONFIG_JSON_PATH):
        try:
            with open(CONFIG_JSON_PATH, "r", encoding="utf-8") as f:
                config_data = json.load(f)
        except Exception:
            pass

    company_name = config_data.get("company_name") or os.getenv("CORPORATE_NAME") or "주식회사 [사명 미설정]"
    base_dir = config_data.get("base_dir") or os.getenv("CORPORATE_BASE_DIR") or r"C:\Enterprise_Assets"
    snapshot_date = config_data.get("snapshot_date") or os.getenv("CORPORATE_SNAPSHOT_DATE") or "2025년 12월 31일"

    return company_name, base_dir, snapshot_date

def save_user_config(company_name: str, base_dir: str, snapshot_date: str = "2025년 12월 31일"):
    """Save user interactive configuration into config.json."""
    config_data = {
        "company_name": company_name.strip(),
        "base_dir": base_dir.strip(),
        "snapshot_date": snapshot_date.strip()
    }
    with open(CONFIG_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)
    return CONFIG_JSON_PATH

def run_interactive_setup():
    """Interactive First-Time Setup Wizard for Company Name and Drive/PC Directory Path."""
    current_company, current_base_dir, current_snapshot = load_user_config()

    print("
=================================================================")
    print("🌐 AX-Work 마스터 프레임워크 초기 환경 설정 (First-Time Setup Wizard)")
    print("=================================================================")
    print("처음 셋팅을 시작합니다. 관리할 기업 사명과 구글드라이브/PC경로를 지정해주세요.
")

    # 1. Google Drive / PC Folder Path Prompt
    print(f"[현재 설정 경로]: {current_base_dir}")
    input_dir = input("1. 구글 드라이브 또는 PC 로컬 저장소 경로를 입력하세요 (엔터 시 현재값 유지): ").strip()
    target_base_dir = input_dir if input_dir else current_base_dir

    # 2. Company Name Prompt
    print(f"
[현재 설정 사명]: {current_company}")
    input_company = input("2. 관리하실 기업의 사명(법인명)을 입력하세요 (예: (주)폭스커넥트) (엔터 시 현재값 유지): ").strip()
    target_company = input_company if input_company else current_company

    # 3. Snapshot Date Prompt
    print(f"
[현재 스냅샷 기준일]: {current_snapshot}")
    input_snapshot = input("3. 보고서 스냅샷 기준일자를 입력하세요 (엔터 시 현재값 유지): ").strip()
    target_snapshot_date = input_snapshot if input_snapshot else current_snapshot

    saved_path = save_user_config(target_company, target_base_dir, target_snapshot_date)

    print("
-----------------------------------------------------------------")
    print("✅ 설정이 성공적으로 저장되었습니다!")
    print(f"   • 법인 사명 : {target_company}")
    print(f"   • 저장 경로 : {target_base_dir}")
    print(f"   • 기준 일자 : {target_snapshot_date}")
    print(f"   • 설정 파일 : {saved_path}")
    print("=================================================================
")

    return target_company, target_base_dir, target_snapshot_date

# Initialize default runtime configurations
DEFAULT_COMPANY_NAME, DEFAULT_BASE_DIR, DEFAULT_SNAPSHOT_DATE = load_user_config()

FONT_FAMILY = "맑은 고딕"

# Color Tokens for Executive UI Styling
COLOR_HEADER_BG = "1E293B"       # Dark Slate Gray
COLOR_SUMMARY_BG = "F1F5F9"      # Light Slate Gray
COLOR_ZEBRA_BG = "F8FAFC"        # Off-white Zebra
COLOR_WHITE_BG = "FFFFFF"
COLOR_BLACK_BORDER = "000000"

# Status Pill Badges
PILL_STYLES = {
    "정상유지": {"bg": "DCFCE7", "fg": "166534"},
    "정상운행": {"bg": "DCFCE7", "fg": "166534"},
    "전대차유지": {"bg": "DBEAFE", "fg": "1E40AF"},
    "묵시적갱신": {"bg": "DBEAFE", "fg": "1E40AF"},
    "유지중": {"bg": "DCFCE7", "fg": "166534"},
    "양수완료": {"bg": "DBEAFE", "fg": "1E40AF"},
    "만기해지": {"bg": "F1F5F9", "fg": "475569"},
    "양도완료": {"bg": "F1F5F9", "fg": "475569"},
    "중도해지": {"bg": "FEE2E2", "fg": "991B1B"},
    "소유권보유": {"bg": "F3E8FF", "fg": "6B21A8"}
}
