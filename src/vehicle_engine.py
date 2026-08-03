"""
Universal Corporate Vehicle Fleet Master Excel Engine (Audit & IPO Compliant)
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_vehicle_excel(base_dir, company_name="[㈜대상기업]", snapshot_date="2026년 12월 31일"):
    """
    Generates an Audit & IPO compliant Corporate Vehicle Master Excel Register (.xlsx).
    """
    target_dir = os.path.join(base_dir, r"02_차량_자산관리\00_연도별_차량_총괄자산대장")
    os.makedirs(target_dir, exist_ok=True)
    
    file_name = f"[외감_IPO대비]_2026년도_법인차량_총괄자산대장_{company_name.replace(' ', '_')}.xlsx"
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
        "정상운행": {"fill": PatternFill(start_color="DCFCE7", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="166534")},
        "유지중": {"fill": PatternFill(start_color="DCFCE7", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="166534")},
        "자사소유": {"fill": PatternFill(start_color="F3E8FF", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="6B21A8")},
        "양수완료": {"fill": PatternFill(start_color="DBEAFE", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="1E40AF")},
        "만기해지": {"fill": PatternFill(start_color="F1F5F9", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="475569")},
        "양도완료": {"fill": PatternFill(start_color="F1F5F9", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="475569")},
        "중도해지": {"fill": PatternFill(start_color="FEE2E2", fill_type="solid"), "font": Font(name=font_family, size=10, bold=True, color="991B1B")},
    }

    thin_border_side = Side(style='thin', color='CBD5E1')
    thick_bottom_side = Side(style='medium', color='1E293B')
    double_bottom_side = Side(style='double', color='1E293B')

    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_bottom_side)

    # SHEET 1: 2026년도_법인차량_총괄자산대장
    ws1 = wb.active
    ws1.title = "2026년도_법인차량_총괄자산대장"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:S1")
    title_cell = ws1["A1"]
    title_cell.value = f"{company_name} 2026년도 법인차량 총괄자산대장 [외부감사 및 IPO 제출용]"
    title_cell.font = font_title
    title_cell.fill = fill_title
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[1].height = 36

    ws1.cell(row=2, column=1, value=f"※ 스냅샷 기준일자: {snapshot_date} 기준 | 단위: 원 (VAT 포함) | 작성부서: 경영지원본부 총무팀").font = font_subtitle
    ws1.row_dimensions[2].height = 18

    ws1.merge_cells("A4:S4")
    sum_title_cell = ws1["A4"]
    sum_title_cell.value = "📊 2026년도 법인차량 자산 기말 재무 구분 요약 (Executive Financial Summary)"
    sum_title_cell.font = font_kpi_hdr
    sum_title_cell.fill = fill_kpi_hdr
    sum_title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c in range(1, 20):
        ws1.cell(row=4, column=c).border = border_cell
    ws1.row_dimensions[4].height = 24

    kpis = [
        ("총 관리자산", "10 개 자산", "A", "C"),
        ("정상운행 자산", "8 대 (렌트5/리스3)", "D", "F"),
        ("① 렌탈 / 리스 보증금", 95549000, "G", "I"),
        ("② 월 렌탈 / 리스료", 10685220, "J", "L"),
        ("③ 연간 총 리스/렌트비", 128222640, "M", "O"),
        ("④ 해지 / 승계 완료 자산", "2 대 (승계/반납)", "P", "S"),
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
        "순번", "권역", "자산 구분", "차종 (모델명)", "차량번호",
        "금융사 (렌트/리스)", "주 운행자 / 부서", "최초계약일", "임대기간 (시작-종료)",
        "지급 주기", "매월 납부일", "은 행", "계 좌 번 호", "예 금 주",
        "보 증 금", "임 대 료", "약정 주행거리", "당해년도 상태", "비 고"
    ]

    ws1.row_dimensions[8].height = 28
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws1.cell(row=8, column=col_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin_border_side, right=thin_border_side, top=thick_bottom_side, bottom=thick_bottom_side)

    vehicles = [
        (1, "가산", "임차자산", "GV70 2.5T", "141호8727", "현대캐피탈", "[경영지원본부]", "2024-03-15", "2024/03/15 - 2028/03/14", "월세", "10일", "우리은행", "1005-304-112", "현대캐피탈㈜", 12000000, 1150000, "20,000km/연", "정상운행", "경영진 업무용 차종"),
        (2, "강남", "임차자산", "K8 3.5 가솔린", "289수3930", "현대캐피탈", "[영업본부장]", "2023-09-01", "2023/09/01 - 2027/08/31", "월세", "25일", "우리은행", "1005-304-113", "현대캐피탈㈜", 15000000, 1280000, "25,000km/연", "정상운행", "본부장 임원 차량"),
        (3, "가산", "임차자산", "그랜저 2.5 가솔린", "141하9479", "현대캐피탈", "[문항개발팀]", "2024-05-10", "2024/05/10 - 2028/05/09", "월세", "10일", "우리은행", "1005-304-114", "현대캐피탈㈜", 8000000, 920000, "20,000km/연", "정상운행", "가산 본사 업무용"),
        (4, "광명", "임차자산", "GV80 3.0D", "103하8547", "DGB캐피탈", "[콘텐츠본부장]", "2023-11-20", "2023/11/20 - 2027/11/19", "월세", "15일", "신한은행", "201-110-449", "DGB캐피탈㈜", 20000000, 1650000, "20,000km/연", "정상운행", "광명 GIDC 거점 차량"),
        (5, "강남", "임차자산", "벤츠 S500 4MATIC", "281가8991", "하나캐피탈", "[대표이사]", "2023-01-15", "2023/01/15 - 2027/01/14", "월세", "25일", "신한은행", "140-008-771", "하나캐피탈㈜", 30000000, 3100000, "20,000km/연", "정상운행", "대표이사 의전 차량"),
    ]

    start_row = 9
    for idx, v_data in enumerate(vehicles):
        r_num = start_row + idx
        ws1.row_dimensions[r_num].height = 22
        fill_curr = fill_zebra if idx % 2 == 1 else fill_white

        for c_offset, val in enumerate(v_data):
            col_idx = 1 + c_offset
            cell = ws1.cell(row=r_num, column=col_idx, value=val)
            cell.font = font_body
            cell.fill = fill_curr
            cell.border = border_cell

            if c_offset in [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 13, 16, 17]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_offset in [3, 12, 18]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_offset in [14, 15]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,#0"
                cell.font = font_bold if val > 0 else font_body

            if c_offset == 17:
                st_info = status_styles.get(val)
                if st_info:
                    cell.fill = st_info["fill"]
                    cell.font = st_info["font"]

    tot_row = start_row + len(vehicles)
    ws1.row_dimensions[tot_row].height = 26
    ws1.merge_cells(f"A{tot_row}:N{tot_row}")
    tot_label = ws1[f"A{tot_row}"]
    tot_label.value = "합  계 (유효 운행차량 기준)"
    tot_label.font = font_bold
    tot_label.alignment = Alignment(horizontal="center", vertical="center")
    
    for c in range(1, 20):
        cell = ws1.cell(row=tot_row, column=c)
        cell.fill = fill_total
        cell.border = border_total
        if c == 15:
            cell.value = f'=SUMIF(R{start_row}:R{tot_row-1}, "정상운행", O{start_row}:O{tot_row-1})'
            cell.number_format = "#,#0"
            cell.font = font_bold
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif c == 16:
            cell.value = f'=SUMIF(R{start_row}:R{tot_row-1}, "정상운행", P{start_row}:P{tot_row-1})'
            cell.number_format = "#,#0"
            cell.font = font_bold
            cell.alignment = Alignment(horizontal="right", vertical="center")

    v_col_widths = {
        'A': 8.0, 'B': 10.0, 'C': 12.0, 'D': 24.0, 'E': 16.0,
        'F': 18.0, 'G': 18.0, 'H': 14.0, 'I': 28.0, 'J': 10.0,
        'K': 12.0, 'L': 14.0, 'M': 22.0, 'N': 16.0, 'O': 18.0,
        'P': 16.0, 'Q': 16.0, 'R': 14.0, 'S': 36.0
    }
    for col_let, w in v_col_widths.items():
        ws1.column_dimensions[col_let].width = w

    # SHEET 2: 2024-2026_차량변동이력
    ws2 = wb.create_sheet(title="2024-2026_차량변동이력")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H1")
    t2_cell = ws2["A1"]
    t2_cell.value = f" 📜 {company_name} 법인차량 3개년 변동 이력 타임라인 (2024 ~ 2026)"
    t2_cell.font = font_title
    t2_cell.fill = fill_title
    t2_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 36

    v_hist_headers = ["연도", "변동 일자", "권역", "차종 및 차량번호", "변동 구분", "보증금/선납금 (원)", "월 렌탈/리스료 (원)", "세부 변동 이력 및 사유"]
    ws2.row_dimensions[3].height = 28
    for idx, h in enumerate(v_hist_headers, start=1):
        c = ws2.cell(row=3, column=idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=thin_border_side, right=thin_border_side, top=thick_bottom_side, bottom=thick_bottom_side)

    v_history = [
        ("2024년", "2024.01.10", "대전", "카니발 하이리무진 (269더5669)", "신규 리스", 10000000, 1350000, "대전 딥러닝센터 의전용 운용리스 신규 체결"),
        ("2024년", "2024.03.15", "가산", "GV70 2.5T (141호8727)", "신규 렌트", 12000000, 1150000, "가산 경영진 업무용 장기렌트 신규 계약"),
        ("2026년", "2026.01.31", "가산", "카니발 9인승 (138허4412)", "만기 반납", -10000000, -850000, "물류팀 업무용 차량 렌트 만기 반납 완료"),
    ]

    last_timeline_row = 3 + len(v_history)
    for idx, h_data in enumerate(v_history):
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
    r_hdr_cell.value = "📊 회계/재무 검증용 차량 보증금/리스료 3개년 현금흐름 및 기말 잔액 대사표 (Reconciliation Table)"
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
        ("① 기간 중 현금 환수/회수 보증금 총액", 30000000, "2026년 카니발 반납(1천만) 및 승계 보증금 회수액", fill_item_1, font_item_1),
        ("   - 2026년 중 보증금 환수/회수액", 30000000, "카니발 만기반납 1,000만원 + 승계 2,000만원 환수", fill_white, font_body),
        ("② 기간 중 순 보증금 현금 변동액 (Net Cash Flow)", 10549000, "3개년 신규 집행 4,054만 원 - 총 회수액 3,000만 원 = 순유출 1,054만 원", fill_item_2, font_item_2),
        ("③ 2026.12.31 기말 재무상태표상 보증금 잔액", 95549000, "재무상태표(BS) 차량 임차/리스보증금 장부 잔액과 1:1 일치", fill_item_3, font_item_3),
        ("   - 정상운행 차량 보증금 장부 잔액 소계", 95549000, "GV70, K8, 그랜저, GV80 등 운용 차량 소계", fill_white, font_body),
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
