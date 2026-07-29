"""
Universal Multi-Tab Real Estate Master Excel Engine (Audit & IPO Compliant)
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from src.config import (
    FONT_FAMILY, COLOR_HEADER_BG, COLOR_SUMMARY_BG, 
    COLOR_ZEBRA_BG, COLOR_WHITE_BG, PILL_STYLES
)

def generate_real_estate_excel(base_dir, company_name="주식회사 폭스에듀", snapshot_date="2025년 12월 31일"):
    """
    Generates an Audit & IPO compliant multi-tab Real Estate Master Excel Register (.xlsx).
    """
    target_dir = os.path.join(base_dir, r"01_부동산_자산관리\00_연도별_부동산_총괄자산대장")
    os.makedirs(target_dir, exist_ok=True)
    
    file_name = f"[외감_IPO대비]_{company_name.replace(' ', '_')}_연도별_부동산_총괄자산대장(2022-2025).xlsx"
    excel_path = os.path.join(target_dir, file_name)
    
    print(f"[RealEstateEngine] Generating Master Real Estate Register at:\n  {excel_path}")
    
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    title_font = Font(name=FONT_FAMILY, size=15, bold=True, color="0F172A")
    subtitle_font = Font(name=FONT_FAMILY, size=9, bold=True, color="475569")
    summary_bar_fill = PatternFill(start_color=COLOR_SUMMARY_BG, end_color=COLOR_SUMMARY_BG, fill_type="solid")
    header_font = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    data_font = Font(name=FONT_FAMILY, size=10, color="0F172A")
    data_bold_font = Font(name=FONT_FAMILY, size=10, bold=True, color="0F172A")
    zebra_fill = PatternFill(start_color=COLOR_ZEBRA_BG, end_color=COLOR_ZEBRA_BG, fill_type="solid")
    white_fill = PatternFill(start_color=COLOR_WHITE_BG, end_color=COLOR_WHITE_BG, fill_type="solid")

    black_border_side = Side(style='thin', color='000000')
    black_thick_top_side = Side(style='medium', color='000000')

    table_header_border = Border(left=black_border_side, right=black_border_side, top=black_thick_top_side, bottom=black_border_side)
    table_data_border = Border(left=black_border_side, right=black_border_side, top=black_border_side, bottom=black_border_side)

    # --- TAB 1: 01_임대차_자산대장 ---
    ws1 = wb.create_sheet(title="01_임대차_자산대장")
    ws1.views.sheetView[0].showGridLines = False
    ws1.column_dimensions['A'].width = 3

    ws1.cell(row=2, column=2, value=f"[{company_name}] 연도별 부동산 임대차 자산대장 (2022~2025)").font = title_font
    ws1.cell(row=3, column=2, value=f"※ 외감/IPO 제출용 (작성 기준일: {snapshot_date}) | 유효 임대차 자산 대장").font = subtitle_font

    ws1.merge_cells("B5:M5")
    bar_cell = ws1.cell(row=5, column=2, value=f"  [작성 기준일: {snapshot_date} 기준]   유지 자산: 총 17건   |   보증금 잔액 합계: ₩ 1,991,000,000 (19억 9,100만원)   |   월 임대료 합계: ₩ 64,324,500 (VAT 별도)")
    bar_cell.font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
    bar_cell.fill = summary_bar_fill
    bar_cell.alignment = Alignment(horizontal='left', vertical='center')
    bar_cell.border = table_data_border

    headers1 = ["연도", "순서", "구분", "물건명 / 주소", "임대인 (명의)", "계약유형", "임대시작일", "임대종료일", "보증금 (원)", "월 임대료 (원)", "계약 상태", "비고 (연도별 변동 및 해지 이력)"]

    ws1.row_dimensions[7].height = 28
    for c_i, text in enumerate(headers1, 2):
        c = ws1.cell(row=7, column=c_i, value=text)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = table_header_border

    lease_data = [
        ("2022", "01", "임대차", "판교 판교동 612", "박동석(김인숙)", "최초임대차", "2021-07-31", "2023-07-30", 40000000, 3500000, "만기해지", "2023.07.31 계약만기 종료"),
        ("2022", "02", "임대차", "대전 골프존 204호,상담실", "㈜골프존뉴딘홀딩스", "승계계약", "2021-08-17", "2023-08-16", 50000000, 8206000, "정상유지", "2023.09.01 계약연장 및 25.05.31 갱신"),
        ("2022", "03", "임대차", "세종 뱅크빌딩 302호", "타이어뱅크(주)", "최초임대차", "2021-10-01", "2023-09-30", 75000000, 8250000, "만기해지", "2023.10.14 만기 퇴거 완료"),
        ("2022", "04", "임대차", "대전 하우스디어반 B동 721호", "정근호", "최초임대차", "2022-04-07", "2024-04-06", 10000000, 750000, "정상유지", "23년 묵시적갱신 -> 24년/25년 연장계약"),
        ("2022", "05", "임대차", "대전 KCC웰츠타워 1202호", "이선영", "최초임대차", "2022-04-11", "2024-04-10", 20000000, 650000, "묵시적갱신", "2023년~2025년 묵시적 갱신 계속 유지"),
        ("2022", "06", "임대차", "대전 스마트시티 2501호", "김동하", "전세", "2022-06-30", "2024-06-29", 1200000000, 0, "묵시적갱신", "전세보증금 12억원 / 24년, 25년 묵시적갱신"),
        ("2022", "07", "임대차", "광명 GIDC 1214_1215호", "하진우", "최초임대차", "2022-08-01", "2024-07-31", 25000000, 2750000, "묵시적갱신", "24년, 25년 묵시적 갱신 계속 유지"),
        ("2022", "08", "임대차", "대전 월평동577 202호", "류지홍", "최초임대차", "2022-12-12", "2024-12-11", 5000000, 770000, "중도해지", "2023.03.31 중도해지 및 퇴거 완료"),
        ("2023", "09", "임대차", "판교 이레빌딩 3층", "화코스텍인터내셔널㈜", "최초임대차", "2023-01-31", "2025-01-30", 100000000, 11000000, "정상유지", "판교 본점 사옥 (28년 2월까지 5년계약)"),
        ("2023", "10", "임대차", "대전 도룡동 385-28", "김순미", "최초임대차", "2023-02-16", "2025-02-15", 100000000, 6050000, "정상유지", "대전 사옥 리저브 (28년 2월까지 5년계약)"),
        ("2023", "11", "임대차", "광명 센트럴자이 1006호", "김영숙", "최초임대차", "2023-02-24", "2025-02-23", 5000000, 770000, "중도해지", "2023.03.31 중도해지 완료"),
        ("2023", "12", "임대차", "대전 스타빌플러스 511호", "은선희", "최초임대차", "2023-04-08", "2025-04-07", 5000000, 638000, "만기해지", "2025.04.07 만기 계약종료"),
        ("2023", "13", "임대차", "대전 골프존 206호", "㈜골프존뉴딘홀딩스", "승계계약", "2023-08-02", "2025-08-01", 50000000, 5428500, "정상유지", "2023.12.01 법인 승계 완료"),
        ("2023", "14", "임대차", "대전 스마트시티 113호", "강남규", "최초임대차", "2023-10-28", "2025-10-27", 30000000, 1100000, "정상유지", "상가 113호 정상 유지"),
        ("2023", "15", "임대차", "대전 스마트시티 115호", "강남규", "최초임대차", "2023-10-28", "2025-10-27", 30000000, 1100000, "정상유지", "상가 115호 정상 유지"),
        ("2023", "16", "임대차", "서초 강남역리가스퀘어 501호", "이봉순", "최초임대차", "2023-11-10", "2025-11-09", 15000000, 1500000, "만기해지", "2024.11.09 만기 계약종료 (서초본점)"),
        ("2024", "17", "임대차", "대전 스마트시티 209호", "이봉순", "최초임대차", "2024-01-31", "2026-01-30", 30000000, 1500000, "정상유지", "상가 209호 정상 유지"),
        ("2024", "18", "임대차", "대전 하우스디어반 C동 711호", "정근호", "최초임대차", "2024-02-19", "2026-02-18", 10000000, 750000, "중도해지", "2024.02.19 중도해지 처리완료"),
        ("2024", "19", "임대차", "대전 스마트시티 604호", "김동하", "최초임대차", "2024-02-26", "2026-02-25", 50000000, 2500000, "정상유지", "스마트시티 604호 정상 유지"),
        ("2024", "20", "임대차", "가산 대륭포스트타워6차 402_403호", "㈜엠씨에스솔루션", "최초임대차", "2024-02-29", "2026-02-27", 46000000, 4600000, "정상유지", "가산 지식산업센터 402_403호"),
        ("2024", "21", "임대차", "대전 하우스디어반 A동 720호", "정근호", "최초임대차", "2024-04-15", "2026-04-14", 10000000, 790000, "정상유지", "하우스디어반 A동 720호"),
        ("2024", "22", "임대차", "강남 도곡로1길23 지하1층~3층", "박재윤(유한회사 청송)", "최초임대차", "2024-11-01", "2026-10-31", 200000000, 14700000, "정상유지", "강남 사옥 통건물 (보증금 2억/월 1,470만/관리비 219만)"),
        ("2024", "23", "임대차", "대전 갑동 388-1", "데이타이음", "최초임대차", "2024-12-23", "2026-12-22", 20000000, 1200000, "정상유지", "갑동 연구소 및 사업장"),
        ("2025", "24", "임대차", "대전 골프존 104호", "㈜골프존뉴딘홀딩스", "최초임대차", "2025-05-14", "2027-05-13", 10000000, 3148750, "만기해지", "2025.05.31 104호 계약종료분리 (204호는 유지)"),
        ("2025", "25", "임대차", "가산 대륭포스트타워6차 1510호", "㈜엠씨에스솔루션", "최초임대차", "2026-04-29", "2028-04-28", 20000000, 2000000, "정상유지", "가산 1510호 신규확장 체결")
    ]

    for row_idx, row_data in enumerate(lease_data, 8):
        ws1.row_dimensions[row_idx].height = 22
        fill_to_use = zebra_fill if row_idx % 2 == 0 else white_fill
        
        for t_idx, val in enumerate(row_data):
            c_i = t_idx + 2
            cell = ws1.cell(row=row_idx, column=c_i)
            cell.border = table_data_border
            cell.fill = fill_to_use

            if c_i in [2, 3]:
                cell.value = str(val)
                cell.number_format = '@'
                cell.font = data_bold_font if c_i == 3 else data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif c_i in [4, 6, 7, 8, 9]:
                cell.value = str(val)
                cell.number_format = '@'
                cell.font = data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif c_i == 5:
                cell.value = str(val)
                cell.font = data_bold_font
                cell.alignment = Alignment(horizontal='left', vertical='center')
            elif c_i in [10, 11]:
                cell.value = val
                cell.number_format = '#,##0'
                cell.font = data_font
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif c_i == 12:
                cell.value = str(val)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if val in PILL_STYLES:
                    cell.fill = PatternFill(start_color=PILL_STYLES[val]["bg"], end_color=PILL_STYLES[val]["bg"], fill_type="solid")
                    cell.font = Font(name=FONT_FAMILY, size=9, bold=True, color=PILL_STYLES[val]["fg"])
            elif c_i == 13:
                cell.value = str(val)
                cell.font = data_font
                cell.alignment = Alignment(horizontal='left', vertical='center')

    # Auto Column Widths
    col_widths = {1: 3, 2: 10, 3: 8, 4: 10, 5: 32, 6: 24, 7: 14, 8: 14, 9: 14, 10: 18, 11: 16, 12: 14, 13: 48}
    for col_idx, width in col_widths.items():
        ws1.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[RealEstateEngine] Real Estate Master Register Generated Successfully!\n")

if __name__ == "__main__":
    from src.config import DEFAULT_BASE_DIR, DEFAULT_COMPANY_NAME
    generate_real_estate_excel(DEFAULT_BASE_DIR, DEFAULT_COMPANY_NAME)
