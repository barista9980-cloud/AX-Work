"""
Universal Corporate Vehicle Fleet Master Excel Engine (Audit & IPO Compliant)
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

def generate_vehicle_excel(base_dir, company_name="주식회사 폭스에듀", snapshot_date="2025년 12월 31일"):
    """
    Generates an Audit & IPO compliant Corporate Vehicle Master Excel Register (.xlsx).
    """
    target_dir = os.path.join(base_dir, r"02_차량_자산관리\00_연도별_차량_총괄자산대장")
    os.makedirs(target_dir, exist_ok=True)
    
    file_name = f"[외감_IPO대비]_{company_name.replace(' ', '_')}_연도별_법인차량_총괄자산대장(2022-2025).xlsx"
    excel_path = os.path.join(target_dir, file_name)
    
    print(f"[VehicleEngine] Generating Corporate Vehicle Register at:\n  {excel_path}")
    
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

    ws = wb.create_sheet(title="01_법인차량_자산대장")
    ws.views.sheetView[0].showGridLines = False
    ws.column_dimensions['A'].width = 3

    ws.cell(row=2, column=2, value=f"[{company_name}] 연도별 법인차량 총괄자산대장 (2022~2025)").font = title_font
    ws.cell(row=3, column=2, value=f"※ 외감/IPO 제출용 (작성 기준일: {snapshot_date}) | 2025.12.31 기준 유효 법인차량 자산 대장").font = subtitle_font

    ws.merge_cells("B5:M5")
    v_bar_cell = ws.cell(row=5, column=2, value=f"  [작성 기준일: {snapshot_date} 기준]   운행 차량: 총 8대 (누적 10대)   |   보증금 잔액 합계: ₩ 95,549,000 (9,554만원)   |   월 렌탈/리스료 합계: ₩ 10,685,220 (VAT 별도)")
    v_bar_cell.font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
    v_bar_cell.fill = summary_bar_fill
    v_bar_cell.alignment = Alignment(horizontal='left', vertical='center')
    v_bar_cell.border = table_data_border

    v_headers = ["연도", "순서", "차종", "차량번호", "금융사 (렌트/리스)", "계약유형", "계약시작일", "만기일자", "보증금 (원)", "월 렌탈/리스료(원)", "계약 상태", "비고 (승계/양도/양수 이력)"]

    ws.row_dimensions[7].height = 28
    for c_i, text in enumerate(v_headers, 2):
        c = ws.cell(row=7, column=c_i, value=text)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = table_header_border

    v_data = [
        ("2022", "01", "GV70", "141호8727", "현대캐피탈", "장기렌트", "2021-12-27", "2026-12-26", 0, 1096150, "양도완료", "2024.08.02 제3자 양도 완료 및 계약종료"),
        ("2022", "02", "K8", "289수3930", "현대캐피탈", "운용리스", "2022-01-04", "2027-01-04", 0, 555600, "정상운행", "법인 대표 임원 전용 리스 차량"),
        ("2022", "03", "그랜저", "141하9479", "현대캐피탈", "장기렌트", "2022-01-04", "2027-01-04", 0, 646580, "정상운행", "영업 및 업무용 장기 렌트 차량"),
        ("2022", "04", "GV80", "103하8547", "DGB(IM캐피탈)", "장기렌트", "2022-02-28", "2027-02-28", 17360000, 1342660, "양수완료", "2025.11.07 법인 승계(양수) 완료"),
        ("2022", "05", "벤츠 S클래스", "281가8991", "하나캐피탈", "운용리스", "2022-03-14", "2027-03-11", 51459000, 3167300, "정상운행", "대표이사 전용 프리미엄 운용리스"),
        ("2022", "06", "카니발", "269더5669", "현대캐피탈", "운용리스", "2022-03-23", "2027-03-23", 0, 984000, "정상운행", "임직원 이동 및 학원 셔틀 리스 차량"),
        ("2022", "07", "스포티지", "167호2430", "우리캐피탈", "장기렌트", "2022-06-24", "2027-06-23", 0, 752500, "정상운행", "연구소 및 현장 지원 렌트 차량"),
        ("2022", "08", "GV80", "197호3290", "하나캐피탈", "장기렌트", "2022-09-28", "2027-09-27", 26730000, 1449580, "정상운행", "경영진 업무용 장기 렌트 차량"),
        ("2022", "09", "GV70", "172하6158", "농협캐피탈", "장기렌트", "2022-11-10", "2027-11-09", 17445000, 1064000, "양도완료", "2024.08.02 제3자 양도 완료 및 계약종료"),
        ("2024", "10", "아우디 A8", "120노2842", "BNK캐피탈", "운용리스", "2024-11-14", "2029-11-13", 0, 1787000, "정상운행", "임원 전용 프리미엄 운용리스 차량")
    ]

    for row_idx, row_data in enumerate(v_data, 8):
        ws.row_dimensions[row_idx].height = 22
        fill_to_use = zebra_fill if row_idx % 2 == 0 else white_fill
        
        for t_idx, val in enumerate(row_data):
            c_i = t_idx + 2
            cell = ws.cell(row=row_idx, column=c_i)
            cell.border = table_data_border
            cell.fill = fill_to_use

            if c_i in [2, 3]:
                cell.value = str(val)
                cell.number_format = '@'
                cell.font = data_bold_font if c_i == 3 else data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif c_i in [4, 5, 6, 7, 8, 9]:
                cell.value = str(val)
                cell.number_format = '@'
                cell.font = data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif c_i in [10, 11]:  # 보증금(10), 월세(11)
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

    v_col_widths = {1: 3, 2: 10, 3: 8, 4: 16, 5: 16, 6: 18, 7: 14, 8: 14, 9: 14, 10: 18, 11: 18, 12: 14, 13: 44}
    for col_idx, width in v_col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[VehicleEngine] Corporate Vehicle Master Register Generated Successfully!\n")

if __name__ == "__main__":
    from src.config import DEFAULT_BASE_DIR, DEFAULT_COMPANY_NAME
    generate_vehicle_excel(DEFAULT_BASE_DIR, DEFAULT_COMPANY_NAME)
