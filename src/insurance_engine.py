"""
Universal Corporate Insurance Master Excel Engine (Audit & IPO Compliant)
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

def generate_insurance_excel(base_dir, company_name="주식회사 폭스에듀", snapshot_date="2025년 12월 31일"):
    """
    Generates an Audit & IPO compliant Corporate Insurance Master Excel Register (.xlsx).
    """
    target_dir = os.path.join(base_dir, r"03_보험_자산관리\00_연도별_보험_총괄자산대장")
    os.makedirs(target_dir, exist_ok=True)
    
    file_name = f"[외감_IPO대비]_{company_name.replace(' ', '_')}_연도별_기업보험_총괄자산대장(2022-2025).xlsx"
    excel_path = os.path.join(target_dir, file_name)
    
    print(f"[InsuranceEngine] Generating Corporate Insurance Register at:\n  {excel_path}")
    
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

    # --- TAB 1: 01_경영인정기보험_자산대장 ---
    ws1 = wb.create_sheet(title="01_경영인정기보험_자산대장")
    ws1.views.sheetView[0].showGridLines = False
    ws1.column_dimensions['A'].width = 3

    ws1.cell(row=2, column=2, value=f"[{company_name}] 연도별 경영인정기보험 자산대장 (2022~2025)").font = title_font
    ws1.cell(row=3, column=2, value=f"※ 외감/IPO 제출용 (작성 기준일: {snapshot_date}) | 경영인정기보험 자산 대장").font = subtitle_font

    ws1.merge_cells("B5:M5")
    i1_bar_cell = ws1.cell(row=5, column=2, value=f"  [작성 기준일: {snapshot_date} 기준]   유지 계약: 총 5건 (대표이사 보장)   |   월 납입 보험료 합계: ₩ 35,764,220 (월 3,576만원)   |   피보험자: 이종탁 대표이사")
    i1_bar_cell.font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
    i1_bar_cell.fill = summary_bar_fill
    i1_bar_cell.alignment = Alignment(horizontal='left', vertical='center')
    i1_bar_cell.border = table_data_border

    i_headers1 = ["연도", "순서", "보험종목", "피보험자/대상", "보험사", "상품명", "증권번호", "보험개시일", "보험만기일", "월 납입액(원)", "계약 상태", "비고 (경영인보장 및 세무사항)"]

    ws1.row_dimensions[7].height = 28
    for c_i, text in enumerate(i_headers1, 2):
        c = ws1.cell(row=7, column=c_i, value=text)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = table_header_border

    i_data1 = [
        ("2022", "01", "경영인정기보험", "이종탁 대표이사", "KDB생명", "(무)VIP경영인정기보험", "0946352030001", "2022-05-31", "2069-03-31", 5347500, "유지중", "대표이사 정기보장 (1차)"),
        ("2022", "02", "경영인정기보험", "이종탁 대표이사", "KDB생명", "(무)VIP경영인정기보험", "0946352030002", "2022-05-31", "2069-03-31", 5347500, "유지중", "대표이사 정기보장 (2차)"),
        ("2023", "03", "경영인정기보험", "이종탁 대표이사", "매트라이프", "Honors 경영인정기보험Plus", "13460791", "2023-12-28", "2069-12-28", 10011540, "유지중", "대표이사 정기보장 (매트라이프)"),
        ("2024", "04", "경영인정기보험", "이종탁 대표이사", "미래에셋생명", "VIP 경영인을 위한 정기보험", "8005286685", "2024-02-07", "2069-02-07", 5016000, "유지중", "대표이사 정기보장 (미래에셋)"),
        ("2024", "05", "경영인정기보험", "이종탁 대표이사", "삼성생명", "삼성 간편경영인정기보험", "41000016223329", "2024-04-22", "2074-04-22", 10041680, "유지중", "대표이사 50년납 정기보장 (삼성생명)")
    ]

    for row_idx, row_data in enumerate(i_data1, 8):
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
            elif c_i in [4, 5, 6, 7, 8, 9, 10]:
                cell.value = str(val)
                cell.number_format = '@'
                cell.font = data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif c_i == 11:
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

    ins_col_widths = {1: 3, 2: 10, 3: 8, 4: 20, 5: 18, 6: 14, 7: 24, 8: 20, 9: 14, 10: 14, 11: 18, 12: 14, 13: 44}
    for col_idx, width in ins_col_widths.items():
        ws1.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(excel_path)
    print(f"[InsuranceEngine] Corporate Insurance Master Register Generated Successfully!\n")

if __name__ == "__main__":
    from src.config import DEFAULT_BASE_DIR, DEFAULT_COMPANY_NAME
    generate_insurance_excel(DEFAULT_BASE_DIR, DEFAULT_COMPANY_NAME)
