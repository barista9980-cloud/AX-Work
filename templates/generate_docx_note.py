import os
import docx
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
    if shd is not None:
        tcPr.remove(shd)
    new_shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(new_shd)

def set_table_borders(table, color="334155", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    existing = tblPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders')
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(borders)

def set_table_width_and_columns(table, col_widths_in_inches):
    total_dxa = sum([int(w * 1440) for w in col_widths_in_inches])
    tblPr = table._tbl.tblPr
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_dxa}" w:type="dxa"/>')
    existing_w = tblPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblW')
    if existing_w is not None:
        tblPr.remove(existing_w)
    tblPr.append(tblW)
    
    for row in table.rows:
        for c_idx, cell in enumerate(row.cells):
            cell_dxa = int(col_widths_in_inches[c_idx] * 1440)
            cell.width = Inches(col_widths_in_inches[c_idx])
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{cell_dxa}" w:type="dxa"/>')
            existing_cw = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcW')
            if existing_cw is not None:
                tcPr.remove(existing_cw)
            tcPr.append(tcW)

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def set_table_header_repeat(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

def strip_keep_next(p):
    p.paragraph_format.keep_with_next = False
    pPr = p._p.get_or_add_pPr()
    kn = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}keepNext')
    if kn is not None:
        pPr.remove(kn)

def format_header_cell(cell, text, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9"):
    cell.text = text
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    strip_keep_next(p)
    
    if not p.runs:
        p.add_run()
    for r in p.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(font_size)
        r.bold = True
        r.font.color.rgb = RGBColor(30, 41, 59)

def format_data_cell(cell, text, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF"):
    cell.text = text
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    strip_keep_next(p)
    
    if not p.runs:
        p.add_run()
    for r in p.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(font_size)
        r.bold = bold
        r.font.color.rgb = RGBColor(15, 23, 42)

def create_contract_note(master_info, docs_list, output_filepath):
    """
    AI Master Generator Engine for Real Estate Contract Management Notes.
    Updated with 1510호 User-Modified Layout:
    - Section 3 Table 2: Dedicated Account Number Row (gridSpan=2) + Full-Width Enlarged '비고 (관리자 참고사항)' Row (gridSpan=4) with numbered items (1. 2. 3.)
    - Section 4 Table 3: Full-Width Enlarged '기타 특약 및 참조사항' Row (gridSpan=2) with numbered items (1. 2. 3.)
    - Total Table Widths: 6.85 inches (17.4 cm) across all 4 tables
    - Clean 2-Page Split & Zero keepNext Black Dots
    """
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    style_normal = doc.styles['Normal']
    style_normal.font.name = '맑은 고딕'
    style_normal.font.size = Pt(10.0)

    t0_cols = [1.40, 2.025, 1.40, 2.025]
    t1_cols = [0.45, 2.95, 0.85, 1.75, 0.85]
    t2_cols = [1.40, 2.025, 1.40, 2.025]
    t3_cols = [2.00, 4.85]

    b_name = master_info.get("building_name", "")
    u_name = master_info.get("unit_name", "")

    # P0 Title: 13pt Bold
    p0 = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
    p0.text = f"[{b_name}] {u_name} 부동산 계약 관리 노트"
    strip_keep_next(p0)
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(2)
    for r in p0.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(13.0)
        r.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)

    # P1 Sub-note: 9pt Gray
    p1 = doc.add_paragraph()
    p1.text = "※ [AI 파싱 완료] 계약서 PDF에서 OCR 정밀 추출된 내용 입니다.\n※ 최초 작성 시 OCR 정밀 추출 결과물 확인 후 검토해 주시면 됩니다."
    strip_keep_next(p1)
    p1.paragraph_format.space_after = Pt(6)
    for r in p1.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(9.0)
        r.font.color.rgb = RGBColor(100, 116, 139)

    # Section 1 Heading: "1. 계약 정보"
    sec1_p = doc.add_paragraph()
    sec1_p.text = "1. 계약 정보"
    strip_keep_next(sec1_p)
    sec1_p.paragraph_format.space_before = Pt(6)
    sec1_p.paragraph_format.space_after = Pt(2)
    for r in sec1_p.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(11.0)
        r.bold = True
        r.font.color.rgb = RGBColor(30, 41, 59)

    # Table 0: 계약 정보
    t0 = doc.add_table(rows=5, cols=4)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t0, color="334155")
    set_table_width_and_columns(t0, t0_cols)

    t0_data = [
        [("건물명 / 호수", True), (f"{b_name} {u_name}", False), ("사용처 (부서/용도)", True), (master_info.get("usage", ""), False)],
        [("계약 유형", True), (master_info.get("contract_type", "최초임대차"), False), ("최초 계약일", True), (master_info.get("initial_date", ""), False)],
        [("임대 기간", True), (master_info.get("period", ""), False), ("매월 납부일", True), (master_info.get("payment_day", ""), False)],
        [("보증금", True), (master_info.get("deposit", ""), False), ("월 임대료", True), (master_info.get("rent", ""), False)],
        [("계약면적 (㎡)", True), (master_info.get("area_m2", ""), False), ("전용면적 / 평수", True), (master_info.get("area_pyung", ""), False)]
    ]

    for r_i, row in enumerate(t0_data):
        set_row_cant_split(t0.rows[r_i])
        for c_i, (val, is_h) in enumerate(row):
            cell = t0.rows[r_i].cells[c_i]
            if is_h:
                format_header_cell(cell, val, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
            else:
                format_data_cell(cell, val, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")

    # Section 2 Heading: "2. 수록 계약서 문서 목록 (총 N건)"
    sec2_p = doc.add_paragraph()
    sec2_p.text = f"2. 수록 계약서 문서 목록 (총 {len(docs_list)}건)"
    strip_keep_next(sec2_p)
    sec2_p.paragraph_format.space_before = Pt(6)
    sec2_p.paragraph_format.space_after = Pt(2)
    for r in sec2_p.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(11.0)
        r.bold = True
        r.font.color.rgb = RGBColor(30, 41, 59)

    # Table 1: Docs List
    t1 = doc.add_table(rows=1, cols=5)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t1, color="334155")
    set_table_width_and_columns(t1, t1_cols)

    set_row_cant_split(t1.rows[0])
    set_table_header_repeat(t1.rows[0])

    t1_headers = ["순서", "문서명 (파일명)", "계약 종류", "계약 당사자 (임대인 → 임차인)", "계약일"]
    for c_i, h in enumerate(t1_headers):
        cell = t1.rows[0].cells[c_i]
        format_header_cell(cell, h, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="E2E8F0")

    for idx, item in enumerate(docs_list, 1):
        r_row = t1.add_row()
        r_cells = r_row.cells
        set_row_cant_split(r_row)

        bg_color = "F8FAFC" if idx % 2 == 0 else "FFFFFF"

        seq_str = f"{idx:02d}"
        display_title = item.get("display_title", "")
        filename_no_ext = item.get("filename_no_ext", "")
        c_type = item.get("contract_type", "")
        parties_str = item.get("parties", "")
        c_date = item.get("contract_date", "")

        format_data_cell(r_cells[0], seq_str, bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex=bg_color)

        cell1 = r_cells[1]
        set_cell_background(cell1, bg_color)
        set_cell_margins(cell1, top=60, bottom=60, left=100, right=100)
        cell1.text = ""
        
        p_title = cell1.paragraphs[0]
        p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_title.paragraph_format.space_before = Pt(1)
        p_title.paragraph_format.space_after = Pt(0)
        strip_keep_next(p_title)
        r_main = p_title.add_run(display_title)
        r_main.font.name = "맑은 고딕"
        r_main.font.size = Pt(10.0)
        r_main.bold = True
        r_main.font.color.rgb = RGBColor(15, 23, 42)

        p_file = cell1.add_paragraph()
        p_file.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_file.paragraph_format.space_before = Pt(0)
        p_file.paragraph_format.space_after = Pt(1)
        strip_keep_next(p_file)
        r_sub = p_file.add_run(f"({filename_no_ext}.pdf)")
        r_sub.font.name = "맑은 고딕"
        r_sub.font.size = Pt(8.5)
        r_sub.font.color.rgb = RGBColor(100, 116, 139)

        format_data_cell(r_cells[2], c_type, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex=bg_color)
        format_data_cell(r_cells[3], parties_str, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex=bg_color)
        format_data_cell(r_cells[4], c_date, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex=bg_color)

    set_table_width_and_columns(t1, t1_cols)

    # MANDATORY CLEAN PAGE BREAK BEFORE SECTION 3
    doc.add_page_break()

    # Section 3 Heading: "3. 임대인 및 납부 계좌 정보"
    sec3_p = doc.add_paragraph()
    sec3_p.text = "3. 임대인 및 납부 계좌 정보"
    strip_keep_next(sec3_p)
    sec3_p.paragraph_format.space_before = Pt(6)
    sec3_p.paragraph_format.space_after = Pt(2)
    for r in sec3_p.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(11.0)
        r.bold = True
        r.font.color.rgb = RGBColor(30, 41, 59)

    # Table 2: 1510호 Style (6 Rows)
    t2 = doc.add_table(rows=6, cols=4)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t2, color="334155")

    # Row 0
    format_header_cell(t2.rows[0].cells[0], "임대인", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t2.rows[0].cells[1], master_info.get("lessor", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")
    format_header_cell(t2.rows[0].cells[2], "임차인", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t2.rows[0].cells[3], master_info.get("lessee", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")

    # Row 1
    format_header_cell(t2.rows[1].cells[0], "임대인 연락처", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t2.rows[1].cells[1], master_info.get("lessor_phone", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")
    format_header_cell(t2.rows[1].cells[2], "관리사무소 연락처", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t2.rows[1].cells[3], master_info.get("mgmt_phone", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")

    # Row 2
    format_header_cell(t2.rows[2].cells[0], "입금 은행", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t2.rows[2].cells[1], master_info.get("bank", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")
    format_header_cell(t2.rows[2].cells[2], "예금주", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t2.rows[2].cells[3], master_info.get("account_holder", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")

    # Row 3: Merge 0-1 (계좌번호) & Merge 2-3 (계좌번호 데이터)
    t2.rows[3].cells[0].merge(t2.rows[3].cells[1])
    t2.rows[3].cells[2].merge(t2.rows[3].cells[3])
    format_header_cell(t2.rows[3].cells[0], "계좌번호", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t2.rows[3].cells[2], master_info.get("account_number", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="FFFFFF")

    # Row 4: Merge ALL 4 cells for "비고" Header
    t2.rows[4].cells[0].merge(t2.rows[4].cells[1]).merge(t2.rows[4].cells[2]).merge(t2.rows[4].cells[3])
    format_header_cell(t2.rows[4].cells[0], "비고 (관리자 참고사항)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")

    # Row 5: Merge ALL 4 cells for "비고" Data
    t2.rows[5].cells[0].merge(t2.rows[5].cells[1]).merge(t2.rows[5].cells[2]).merge(t2.rows[5].cells[3])
    remarks_text = master_info.get("remarks_full", master_info.get("remarks", "1. 특이사항 없음"))
    format_data_cell(t2.rows[5].cells[0], remarks_text, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex="FFFFFF")

    set_table_width_and_columns(t2, t2_cols)
    for r in t2.rows:
        set_row_cant_split(r)

    # Section 4 Heading: "4. 계약 변동 이력 및 특이사항"
    sec4_p = doc.add_paragraph()
    sec4_p.text = "4. 계약 변동 이력 및 특이사항"
    strip_keep_next(sec4_p)
    sec4_p.paragraph_format.space_before = Pt(6)
    sec4_p.paragraph_format.space_after = Pt(2)
    for r in sec4_p.runs:
        r.font.name = "맑은 고딕"
        r.font.size = Pt(11.0)
        r.bold = True
        r.font.color.rgb = RGBColor(30, 41, 59)

    # Table 3: 1510호 Style (5 Rows)
    t3 = doc.add_table(rows=5, cols=2)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t3, color="334155")

    # Row 0
    format_header_cell(t3.rows[0].cells[0], "계약 변동 / 전대차 / 승계 이력", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t3.rows[0].cells[1], master_info.get("history_text", "특이 변동이력 없음 (최초 계약 유지 중)"), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex="FFFFFF")

    # Row 1
    format_header_cell(t3.rows[1].cells[0], "계약 연장 / 묵시적 갱신 이력", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t3.rows[1].cells[1], master_info.get("renewal_text", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex="FFFFFF")

    # Row 2
    format_header_cell(t3.rows[2].cells[0], "중도해지 / 퇴거 예정 메모", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")
    format_data_cell(t3.rows[2].cells[1], master_info.get("termination_text", ""), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex="FFFFFF")

    # Row 3: Merge both cells for Header
    t3.rows[3].cells[0].merge(t3.rows[3].cells[1])
    format_header_cell(t3.rows[3].cells[0], "기타 특약 및 참조사항", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex="F1F5F9")

    # Row 4: Merge both cells for Data
    t3.rows[4].cells[0].merge(t3.rows[4].cells[1])
    special_terms_text = master_info.get("special_terms_full", master_info.get("special_notes", "1. 계약일 현재 등기부등본 확인 후 대상부동산의 권리 및 시설상태의 계약으로 한다.\n2. 부가세 및 관리비는 별도로 한다."))
    format_data_cell(t3.rows[4].cells[0], special_terms_text, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex="FFFFFF")

    set_table_width_and_columns(t3, t3_cols)
    for r in t3.rows:
        set_row_cant_split(r)

    # Extra safety check: strip keep_with_next from EVERY paragraph
    for p in doc.paragraphs:
        strip_keep_next(p)

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for cp in c.paragraphs:
                    strip_keep_next(cp)

    doc.save(output_filepath)
    print("Master Contract Note Template saved successfully:", output_filepath)
