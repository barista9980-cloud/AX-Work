import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
config_p = os.path.join(REPO_DIR, r"src\config.py")

config_content = """\"\"\"
Universal Configuration for Corporate Asset Management Framework
Supports Environment Variables and CLI Overrides for Multi-Entity Deployment.
\"\"\"
import os
import sys

# Dynamic Corporate Metadata (Overridden via CLI --company / --base-dir or ENV variables)
DEFAULT_COMPANY_NAME = os.getenv("CORPORATE_NAME", "주식회사 폭스에듀")
DEFAULT_BASE_DIR = os.getenv("CORPORATE_BASE_DIR", r"G:\\내 드라이브\\[FoxConnect]\\[총무]업무")
DEFAULT_SNAPSHOT_DATE = os.getenv("CORPORATE_SNAPSHOT_DATE", "2025년 12월 31일")

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
"""

with open(config_p, "w", encoding="utf-8") as f:
    f.write(config_content)

print("Updated src/config.py with Dynamic Environment Variable Support!")
