import os
import docx
from templates.generate_docx_note import create_contract_note

desktop_path = r"C:\Users\User\OneDrive\바탕 화면\부동산_계약관리노트_402_403호.docx"

master_info = {
    "building_name": "01_가산_대륭포스트타워6차",
    "unit_name": "402_403호",
    "usage": "가산 사무실(콘텐츠본부, 전략본부)",
    "contract_type": "최초임대차",
    "initial_date": "2024-02-29",
    "period": "2024-02-29 ~ ",
    "payment_day": "",
    "deposit": "",
    "rent": "",
    "area_m2": "",
    "area_pyung": "",
    "lessor": "㈜엠씨에스솔루션",
    "lessee": "㈜폭스에듀(폭스커넥트 법인)",
    "lessor_phone": "",
    "mgmt_phone": "",
    "bank": "",
    "account_holder": "㈜엠씨에스솔루션",
    "account_number": "",
    "remarks": "특이사항 없음",
    "history_text": "2025-07-01[전대차] ㈜폭스에듀 → 정선혜",
    "renewal_text": "",
    "termination_text": "",
    "special_notes": ""
}

docs_list = [
    {
        "display_title": "402_403호 최초임대차계약서",
        "filename_no_ext": "가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229)",
        "contract_type": "최초임대차",
        "parties": "㈜엠씨에스솔루션 → ㈜폭스에듀",
        "contract_date": "2024-02-29"
    },
    {
        "display_title": "403호 전대차계약서",
        "filename_no_ext": "가산_대륭포스트타워6차_403호_02_전대차_[㈜폭스에듀-정선혜]_(250701)",
        "contract_type": "전대차",
        "parties": "㈜폭스에듀 → 정선혜",
        "contract_date": "2025-07-01"
    }
]

create_contract_note(master_info, docs_list, desktop_path)
print("Updated desktop reference master note template:", desktop_path)
