"""
Universal Corporate Insurance Master Excel Engine (Audit & IPO Compliant)
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_insurance_excel(base_dir, company_name="[㈜대상기업]", snapshot_date="2026년 12월 31일"):
    """
    Generates an Audit & IPO compliant Corporate Insurance Master Excel Register (.xlsx).
    """
    target_dir = os.path.join(base_dir, r"03_보험_자산관리\00_연도별_보험_총괄자산대장")
    os.makedirs(target_dir, exist_ok=True)
    
    file_name = f"[외감_IPO대비]_2026년도_기업보험_총괄자산대장_{company_name.replace(' ', '_')}.xlsx"
    excel_path = os.path.join(target_dir, file_name)
    
    wb = openpyxl.Workbook()
    
    font_family = "맑은 고딕"
    font_title = Font(name=font_family, size=15, bold=True, color="FFFFFF")
    font_subtitle = Font(name=font_family, size=10, italic=True, color="475569")
    font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    font_kpi_hdr = Font(name=font_family, size=11, bold=True, color="1E293B")
    font_kpi_label = Font(name=font_family, size=10, bold=False, color="0F172A")
    font_kpi_val = Font(name=font_family, size=10, bold=True, color="0F172A")
    font_body = Font(name=font_family, size=10, color="0F172A")
    font_bold = Font(name=font_family, size=10, bold=True, color="0F172A")

    fill_title = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_kpi_hdr = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_kpi_label = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    fill_rec_hdr = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_item_1 = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    font_item_1 = Font(name=font_family, size=10, bold=True, color="166534")

    fill_item_2 = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_item_2 = Font(name=font_family, size=10, bold=True, color="92400E")

    fill_item_3 = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    font_item_3 = Font(name=font_family, size=10, bold=True, color="1E40AF")

    status_styles = {
        "유지중": {"fill": PatternFill(start_color="DCFCE7", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="166534")},
        "정상유지": {"fill": PatternFill(start_color="DCFCE7", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="166534")},
        "만기해지": {"fill": PatternFill(start_color="F1F5F9", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="475569")},
        "해지완료": {"fill": PatternFill(start_color="FEE2E2", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="991B1B")},
    }

    thin_border_side = Side(style='thin', color='CBD5E1')
    thick_bottom_side = Side(style='medium', color='1E293B')
    double_bottom_side = Side(style='double', color='1E293B')

    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_bottom_side)

    # SHEET 1: 2026년도_기업보험_총괄자산대장
    ws1 = wb.active
    ws1.title = "2026년도_기업보험_총괄자산대장"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:Q1")
    title_cell = ws1["A1"]
    title_cell.value = f"{company_name} 2026년도 기업보험 총괄자산대장 [외부감사 및 IPO 제출용]"
    title_cell.font = font_title
    title_cell.fill = fill_title
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[1].height = 36

    ws1.cell(row=2, column=1, value=f"※ 스냅샷 기준일자: {snapshot_date} 기준 | 단위: 원 (VAT 포함) | 작성부서: 경영지원본부 총무팀").font = font_subtitle
    ws1.row_dimensions[2].height = 18

    ws1.merge_cells("A4:Q4")
    sum_title_cell = ws1["A4"]
    sum_title_cell.value = "📊 2026년도 기업보험 자산 기말 재무 구분 요약 (Executive Financial Summary)"
    sum_title_cell.font = font_kpi_hdr
    sum_title_cell.fill = fill_kpi_hdr
    sum_title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c in range(1, 18):
        ws1.cell(row=4, column=c).border = border_cell
    ws1.row_dimensions[4].height = 24

    kpis = [
        ("총 관리자산", "14 건 (유지12/해지2)", "A", "C"),
        ("정상유지 계약", "12 건 (임원5/화재4/배상3)", "D", "F"),
        ("① 월 납입 총 보험료", 35764220, "G", "I"),
        ("② 연간 총 납입 보험료", 429170640, "J", "L"),
        ("③ 총 가입 보장 한도", 15000000000, "M", "O"),
        ("④ 해지 / 만기 완료 계약", "2 건 (해약환급 완료)", "P", "Q"),
    ]

    for label, val, c_start, c_end in kpis:
        ws1.merge_cells(f"{c_start}5:{c_end}5")
        l_cell = ws1[f"{c_start}5"]
        l_cell.value = label
        l_cell.font = font_kpi_label
        l_cell.fill = fill_kpi_label
        l_cell.alignment = Alignment(horizontal="center", vertical="center")

        start_col = openpyxl.utils.column_index_from_string(c_start)
        end_col = openpyxl.utils.column_index_from_string(c_end)
        for col_idx in range(start_col, end_col + 1):
            ws1.cell(row=5, column=col_idx).border = border_cell

        ws1.merge_cells(f"{c_start}6:{c_end}6")
        v_cell = ws1[f"{c_start}6"]
        v_cell.value = val
        v_cell.font = font_kpi_val
        v_cell.alignment = Alignment(horizontal="center", vertical="center")
        if isinstance(val, (int, float)):
            v_cell.number_format = "#,#0"

        for col_idx in range(start_col, end_col + 1):
            ws1.cell(row=6, column=col_idx).border = border_cell

    ws1.row_dimensions[5].height = 20
    ws1.row_dimensions[6].height = 24

    headers = [
        "순번", "보험 종목 구분", "증권 번호 (Policy No.)", "보험사명", "계약자 (법인명)",
        "피보험자 / 보장 대상", "최초 가입일", "보험 기간 (시작-종료)", "지급 주기",
        "매월 납부일", "은 행", "계 좌 번 호", "예 금 주", "보장 한도 (원)",
        "월 보험료 (원)", "당해년도 상태", "비 고 (수익자 / 특약사항)"
    ]

    ws1.row_dimensions[8].height = 28
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws1.cell(row=8, column=col_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin_border_side, right=thin_border_side, top=thick_bottom_side, bottom=thick_bottom_side)

    policies = [
        (1, "경영인정기보험", "POL-2023-08912", "삼성생명", "[㈜대상기업]", "[이종탁 대표이사]", "2023-04-15", "2023/04/15 - 2043/04/14", "월세", "15일", "우리은행", "1005-304-638857", "삼성생명보험㈜", 3000000000, 8500000, "유지중", "수익자: [㈜대상기업] / 대표이사 유해 보장"),
        (2, "경영인정기보험", "POL-2023-08913", "한화생명", "[㈜대상기업]", "[김철수 부사장]", "2023-05-01", "2023/05/01 - 2043/04/30", "월세", "10일", "우리은행", "1005-304-638857", "한화생명보험㈜", 2000000000, 6200000, "유지중", "수익자: [㈜대상기업] / 핵심 임원 보장"),
        (3, "경영인정기보험", "POL-2024-01201", "DB생명", "[㈜대상기업]", "[박민우 본부장]", "2024-02-10", "2024/02/10 - 2044/02/09", "월세", "10일", "우리은행", "1005-304-638857", "DB생명보험㈜", 1500000000, 4800000, "유지중", "수익자: [㈜대상기업] / 임원 경영 안정"),
        (4, "화재/재산보험", "FIRE-2024-0012", "DB손해보험", "[㈜대상기업]", "[가산 대륭포스트타워 402호]", "2024-03-01", "2024/03/01 - 2027/02/28", "연납", "3/1일", "우리은행", "1005-304-638857", "DB손해보험㈜", 2000000000, 450000, "정상유지", "문항개발실 재산 종합 화재보험"),
    ]

    start_row = 9
    for idx, p_data in enumerate(policies):
        r_num = start_row + idx
        ws1.row_dimensions[r_num].height = 22
        fill_curr = fill_zebra if idx % 2 == 1 else fill_white

        for c_offset, val in enumerate(p_data):
            col_idx = 1 + c_offset
            cell = ws1.cell(row=r_num, column=col_idx, value=val)
            cell.font = font_body
            cell.fill = fill_curr
            cell.border = border_cell

            if c_offset in [0, 1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 15]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_offset in [5, 16]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_offset in [13, 14]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,#0"
                cell.font = font_bold if val > 0 else font_body

            if c_offset == 15:
                st_info = status_styles.get(val)
                if st_info:
                    cell.fill = st_info["fill"]
                    cell.font = st_info["font"]

    tot_row = start_row + len(policies)
    ws1.row_dimensions[tot_row].height = 26
    ws1.merge_cells(f"A{tot_row}:M{tot_row}")
    tot_label = ws1[f"A{tot_row}"]
    tot_label.value = "합  계 (유효 유지계약 기준)"
    tot_label.font = font_bold
    tot_label.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, 18):
        cell = ws1.cell(row=tot_row, column=c)
        cell.fill = fill_total
        cell.border = border_total
        if c == 14:
            cell.value = f'=SUMIF(P{start_row}:P{tot_row-1}, "유지중", N{start_row}:N{tot_row-1})'
            cell.number_format = "#,#0"
            cell.font = font_bold
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif c == 15:
            cell.value = f'=SUMIF(P{start_row}:P{tot_row-1}, "유지중", O{start_row}:O{tot_row-1})'
            cell.number_format = "#,#0"
            cell.font = font_bold
            cell.alignment = Alignment(horizontal="right", vertical="center")

    i_col_widths = {
        'A': 8.0, 'B': 18.0, 'C': 22.0, 'D': 16.0, 'E': 16.0,
        'F': 28.0, 'G': 14.0, 'H': 28.0, 'I': 10.0, 'J': 12.0,
        'K': 14.0, 'L': 22.0, 'M': 18.0, 'N': 20.0, 'O': 16.0,
        'P': 14.0, 'Q': 42.0
    }
    for col_let, w in i_col_widths.items():
        ws1.column_dimensions[col_let].width = w

    # SHEET 2: 2024-2026_보험변동이력
    ws2 = wb.create_sheet(title="2024-2026_보험변동이력")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H1")
    t2_cell = ws2["A1"]
    t2_cell.value = f" 📜 {company_name} 기업보험 3개년 변동 이력 타임라인 (2024 ~ 2026)"
    t2_cell.font = font_title
    t2_cell.fill = fill_title
    t2_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 36

    i_hist_headers = ["연도", "변동 일자", "보험 종목", "증권번호 및 보장 대상", "변동 구분", "월/연 납입액 (원)", "환급금/수령액 (원)", "세부 변동 이력 및 배서 내역"]
    ws2.row_dimensions[3].height = 28
    for idx, h in enumerate(i_hist_headers, start=1):
        c = ws2.cell(row=3, column=idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=thin_border_side, right=thin_border_side, top=thick_bottom_side, bottom=thick_bottom_side)

    i_history = [
        ("2024년", "2024.01.01", "배상책임", "전 사업장 시설배상책임 (LIAB-2024-0991)", "신규 가입", 3000000, 0, "2024년도 전 사업장 시설소유자 배상책임 신규 체결"),
        ("2024년", "2024.02.10", "경영인정기", "DB생명 박민우 본부장 (POL-2024-01201)", "신규 가입", 4800000, 0, "핵심 임원 경영 안정용 경영인정기보험 신규 계약"),
        ("2025년", "2025.02.28", "경영인정기", "흥국생명 전임 임원 (POL-2021-00101)", "만기 해지", -4000000, 45000000, "전임 임원 퇴임에 따른 계약 해지 및 해약환급금 4,500만원 입금 완료"),
    ]

    last_timeline_row = 3 + len(i_history)
    for idx, h_data in enumerate(i_history):
        r_num = 4 + idx
        ws2.row_dimensions[r_num].height = 22
        fill_curr = fill_zebra if idx % 2 == 1 else fill_white

        for c_off, val in enumerate(h_data):
            col = 1 + c_off
            cell = ws2.cell(row=r_num, column=col, value=val)
            cell.font = font_body
            cell.fill = fill_curr
            cell.border = border_cell

            if c_off in [0, 1, 2, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_off == 3 or c_off == 7:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_off in [5, 6]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,#0"

    rec_start_row = last_timeline_row + 2
    ws2.merge_cells(f"A{rec_start_row}:H{rec_start_row}")
    r_hdr_cell = ws2[f"A{rec_start_row}"]
    r_hdr_cell.value = "📊 회계/재무 검증용 기업보험 3개년 현금흐름 및 해약환급금 대사표 (Reconciliation Table)"
    r_hdr_cell.font = font_kpi_hdr
    r_hdr_cell.fill = fill_rec_hdr
    r_hdr_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c in range(1, 9):
        ws2.cell(row=rec_start_row, column=c).border = border_cell
    ws2.row_dimensions[rec_start_row].height = 26

    rec_col_hdr_row = rec_start_row + 1
    ws2.merge_cells(f"A{rec_col_hdr_row}:D{rec_col_hdr_row}")
    ws2[f"A{rec_col_hdr_row}"].value = "회계 대사 구분 항목"
    ws2[f"A{rec_col_hdr_row}"].alignment = Alignment(horizontal="center", vertical="center")

    ws2.merge_cells(f"E{rec_col_hdr_row}:F{rec_col_hdr_row}")
    ws2[f"E{rec_col_hdr_row}"].value = "금액 (원)"
    ws2[f"E{rec_col_hdr_row}"].alignment = Alignment(horizontal="center", vertical="center")

    ws2.merge_cells(f"G{rec_col_hdr_row}:H{rec_col_hdr_row}")
    ws2[f"G{rec_col_hdr_row}"].value = "회계장부(BS) 및 현금흐름 대사 설명"
    ws2[f"G{rec_col_hdr_row}"].alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, 9):
        cell = ws2.cell(row=rec_col_hdr_row, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_cell
    ws2.row_dimensions[rec_col_hdr_row].height = 24

    reconciliations = [
        ("① 기간 중 해약환급금 현금 수령 총액", 45000000, "2025년 흥국생명 경영인정기보험 만기 해약환급금 통장 입금액", fill_item_1, font_item_1),
        ("   - 2025년 중 해약환급금 수령액", 45000000, "흥국생명 경영인정기보험 해지환급금 4,500만 원 수령", fill_white, font_body),
        ("② 기간 중 순 보험료 현금 납입 총액 (Net Cash Outflow)", 429170640, "2026.12.31 기말 기준 12개 유지 계약 연간 총 납입 보험료", fill_item_2, font_item_2),
        ("③ 2026.12.31 기말 재무상태표상 장급비용 / 환급금 장부 잔액", 291000000, "경영인정기보험 해약환급금 장부 가치 잔액 (5개 경영인정기보험 합계)", fill_item_3, font_item_3),
        ("   - 경영인정기보험 해약환급 장부 잔액 소계", 291000000, "삼성생명, 한화생명, DB생명, 교보생명, KB라이프 5개 소계", fill_white, font_body),
    ]

    r_item_base_row = rec_col_hdr_row + 1
    for idx, r_data in enumerate(reconciliations):
        r_num = r_item_base_row + idx
        ws2.row_dimensions[r_num].height = 22
        item_title, item_amt, item_desc, item_fill, item_font = r_data

        ws2.merge_cells(f"A{r_num}:D{r_num}")
        c_item = ws2[f"A{r_num}"]
        c_item.value = item_title
        c_item.font = item_font
        c_item.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws2.merge_cells(f"E{r_num}:F{r_num}")
        c_amt = ws2[f"E{r_num}"]
        c_amt.value = item_amt
        c_amt.font = item_font
        c_amt.alignment = Alignment(horizontal="right", vertical="center")
        c_amt.number_format = "#,#0"

        ws2.merge_cells(f"G{r_num}:H{r_num}")
        c_desc = ws2[f"G{r_num}"]
        c_desc.value = item_desc
        c_desc.font = item_font if item_fill != fill_white else font_body
        c_desc.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        for col_idx in range(1, 9):
            ws2.cell(row=r_num, column=col_idx).fill = item_fill
            ws2.cell(row=r_num, column=col_idx).border = border_cell

    s2_col_widths = {
        'A': 10.0, 'B': 14.0, 'C': 10.0, 'D': 35.0,
        'E': 14.0, 'F': 20.0, 'G': 24.0, 'H': 68.0
    }
    for col_let, w in s2_col_widths.items():
        ws2.column_dimensions[col_let].width = w

    wb.save(excel_path)
    return excel_path
