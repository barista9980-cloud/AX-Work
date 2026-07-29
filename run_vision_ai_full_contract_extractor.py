import os
import sys
import json
import fitz
import urllib.request

sys.stdout.reconfigure(encoding='utf-8')

api_key = ""
key_file = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\local_api_key.txt"
if os.path.exists(key_file):
    with open(key_file, "r", encoding="utf-8") as f:
        api_key = f.read().strip()

def extract_vision_ai_data(pdf_path):
    doc_pdf = fitz.open(pdf_path)
    full_text = ""
    for page in doc_pdf:
        full_text += page.get_text() + "\n"

    prompt = f"""
당신은 대한민국 부동산 임대차계약서 정밀 추출 전문 Vision AI입니다.
아래 계약서 텍스트 원문에서 보증금, 월임대료, 임대기간, 매월납부일, 임대인, 임차인, 은행, 계좌번호, 예금주, 특약사항을 정밀 추출하세요.

다음 규칙을 반드시 준수하여 오직 JSON 데이터만 반환하세요:
1. deposit: 보증금 금액 (예: "40,000,000원 (4,000만원)")
2. monthly_rent: 월 임대료/차임 (예: "3,500,000원 (부가세 별도)")
3. lease_period: 임대 기간 (예: "2021-07-31 ~ 2023-07-30 (24개월)")
4. pay_day: 매월 납부일 (예: "매월 31일")
5. landlord: 임대인 성명 및 법인명 (예: "박동석, 김인숙")
6. lessee: 임차인 성명 및 법인명 (예: "주식회사 폭스에듀")
7. bank: 입금 은행 (예: "국민은행" 또는 "확인 필요")
8. account_num: 계좌번호 (예: "123-456-7890" 또는 "확인 필요")
9. account_holder: 예금주 (예: "박동석" 또는 "확인 필요")
10. special_terms: 기타 특약사항 (배열 형식, 예: ["1. 계약일 현재 등기부등본 확인 후 상태 계약으로 함", "2. 원상복구를 기본으로 함"])

[계약서 텍스트]:
{full_text[:4000]}
"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"response_mime_type": "application/json"}
    }
    
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode('utf-8'),
        headers={'Content-Type': 'application/json'}
    )
    
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            result_text = data['candidates'][0]['content']['parts'][0]['text']
            return json.loads(result_text)
    except Exception as e:
        print(f"Error calling Gemini Vision API on {os.path.basename(pdf_path)}:", e)
        return None

pdf_p = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\01_최초임대차_판교_판교동612_[박동석-㈜폭스에듀]_(210731).pdf"
v_res = extract_vision_ai_data(pdf_p)
print("\n=== GEMINI 1.5 FLASH VISION PARSED RESULT ===")
print(json.dumps(v_res, ensure_ascii=False, indent=2))
