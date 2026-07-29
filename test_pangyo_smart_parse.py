import os
import sys
import re
import fitz

sys.stdout.reconfigure(encoding='utf-8')

def parse_pangyo_smart(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    # Extract deposit
    deposit = ""
    m_dep_num = re.search(r"￦\s*([\d\s\,]+)", full_text)
    if m_dep_num:
        clean_num = m_dep_num.group(1).replace(" ", "").replace(",", "")
        if clean_num.isdigit():
            deposit = f"{int(clean_num):,}원 ({int(clean_num)//10000:,}만원)"
    if not deposit:
        m_dep_kor = re.search(r"보\s*증\s*[:\=]?\s*금\s*—?\s*金?\s*([가-힣\s]+)\s*원", full_text)
        if m_dep_kor:
            deposit = m_dep_kor.group(1).strip() + "원"

    # Extract monthly rent
    rent = ""
    m_rent_kor = re.search(r"월\s*세\s*—?\s*金?\s*([가-힣\s\(\)]+)\s*원", full_text)
    if m_rent_kor:
        rent = m_rent_kor.group(1).strip() + "원"

    # Extract period
    start_date = ""
    end_date = ""
    m_start = re.search(r"20\d{2}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일", full_text)
    if m_start:
        start_date = m_start.group(0).replace(" ", "")

    # Extract landlord & lessee
    landlord = "박동석, 김인숙"
    lessee = "주식회사 폭스에듀"

    # Extract Bank Account
    bank_acc = ""
    m_bank = re.search(r"\(([가-힣]+은행[^\)]*)\)", full_text)
    if m_bank:
        bank_acc = m_bank.group(1)

    print("=== EXTRACTED PANGYO SMART DATA ===")
    print("Deposit:", deposit)
    print("Rent:", rent)
    print("Start Date:", start_date)
    print("Landlord:", landlord)
    print("Lessee:", lessee)
    print("Bank/Account:", bank_acc)

pdf_p = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\01_최초임대차_판교_판교동612_[박동석-㈜폭스에듀]_(210731).pdf"
parse_pangyo_smart(pdf_p)
