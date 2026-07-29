import os
import sys
import re
import fitz

sys.stdout.reconfigure(encoding='utf-8')

def parse_korean_amount(text):
    """
    Parses Korean monetary amounts like "사천만원", "삼백오십만원", "40,000,000", "46,000,000".
    """
    if not text:
        return ""
    
    # Clean text
    clean_t = text.replace(" ", "").replace(",", "").replace("—", "").replace("金", "")
    
    # Try finding digits like 40000000
    m_digits = re.search(r"(\d{6,10})", clean_t)
    if m_digits:
        val = int(m_digits.group(1))
        if val >= 10000:
            man = val // 10000
            rem = val % 10000
            if rem > 0:
                return f"{val:,}원 ({man:,}만 {rem:,}원)"
            else:
                return f"{val:,}원 ({man:,}만원)"

    # Korean text mapping
    kor_map = {
        "일천만": 10000000, "이천만": 20000000, "삼천만": 30000000, "사천만": 40000000,
        "오천만": 50000000, "육천만": 60000000, "칠천만": 70000000, "팔천만": 80000000, "구천만": 90000000,
        "일억": 100000000, "이억": 200000000, "삼억": 300000000, "사억": 400000000, "오억": 500000000,
        "백만": 1000000, "이백만": 2000000, "삼백만": 3000000, "사백만": 4000000, "오백만": 5000000,
        "삼백오십만": 3500000, "이백오십만": 2500000, "백오십만": 1500000, "사백오십만": 4500000, "오백오십만": 5500000
    }
    for k_str, val in kor_map.items():
        if k_str in clean_t:
            man = val // 10000
            vat_str = " (VAT 별도)" if "부가세" in clean_t or "부가세별도" in text else ""
            return f"{val:,}원 ({man:,}만원){vat_str}"

    return text.strip()

def parse_full_contract_details(pdf_path):
    doc_pdf = fitz.open(pdf_path)
    full_text = ""
    for page in doc_pdf:
        full_text += page.get_text() + "\n"

    # Deposit
    deposit = ""
    m_dep = re.search(r"보\s*증\s*[:\=]?\s*금[^\n]*\n?([^\n]+)", full_text)
    if m_dep:
        deposit = parse_korean_amount(m_dep.group(1))

    # Monthly Rent
    rent = ""
    m_rent = re.search(r"월\s*세[^\n]*\n?([^\n]+)", full_text)
    if not m_rent:
        m_rent = re.search(r"차\s*임[^\n]*\n?([^\n]+)", full_text)
    if m_rent:
        rent = parse_korean_amount(m_rent.group(1))

    # Dates
    start_d = ""
    end_d = ""
    m_dates = re.findall(r"20\d{2}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일", full_text)
    if len(m_dates) >= 2:
        start_d = m_dates[0].replace(" ", "")
        end_d = m_dates[1].replace(" ", "")
    elif len(m_dates) == 1:
        start_d = m_dates[0].replace(" ", "")

    # Pay day
    pay_day = ""
    m_pay = re.search(r"매월\s*(\d{1,2}\s*일)", full_text)
    if m_pay:
        pay_day = f"매월 {m_pay.group(1).strip()}"

    # Landlord & Lessee
    landlord = ""
    lessee = ""
    m_landlord = re.search(r"임\s*대\s*인[^\n]*성\s*명\s*[:\=]?\s*([^\n]+)", full_text)
    if m_landlord:
        landlord = m_landlord.group(1).strip()
    
    m_lessee = re.search(r"임\s*차\s*인[^\n]*성\s*명\s*[:\=]?\s*([^\n]+)", full_text)
    if m_lessee:
        lessee = m_lessee.group(1).strip()

    # Bank
    bank = ""
    m_bank = re.search(r"(국민|신한|우리|하나|기업|농협|카카오|케이|수협|대구|부산|경남|광주|전북|우체국)\s*은행", full_text)
    if m_bank:
        bank = m_bank.group(0)

    # Account Number
    acc_num = ""
    m_acc = re.search(r"(\d{3,6}[\-\s]\d{2,6}[\-\s]\d{3,8})", full_text)
    if m_acc:
        acc_num = m_acc.group(1).strip()

    return {
        "deposit": deposit,
        "monthly_rent": rent,
        "start_date": start_d,
        "end_date": end_d,
        "pay_day": pay_day,
        "landlord": landlord,
        "lessee": lessee,
        "bank": bank,
        "account_num": acc_num
    }

pdf_p = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\01_최초임대차_판교_판교동612_[박동석-㈜폭스에듀]_(210731).pdf"
res = parse_full_contract_details(pdf_p)
print("=== ROBUST KOREAN CONTRACT PARSER RESULT FOR PANGYO ===")
for k, v in res.items():
    print(f"  {k}: {v}")
