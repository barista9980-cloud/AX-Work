import os
import sys
import json
import fitz
import urllib.request
import urllib.parse

sys.stdout.reconfigure(encoding='utf-8')

# Retrieve API Key from Environment or Config
api_key = os.environ.get("GEMINI_API_KEY", "")

# Fallback: check if stored in antigravity environment
if not api_key:
    # Try reading from config file or user env
    api_key_file = os.path.expanduser(r"~\.gemini\gemini_api_key.txt")
    if os.path.exists(api_key_file):
        with open(api_key_file, "r", encoding="utf-8") as f:
            api_key = f.read().strip()

print("Gemini API Key Available:", bool(api_key))

def extract_with_gemini_api(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    prompt = f"""
다음은 대한민국 부동산 임대차계약서 텍스트입니다.
텍스트에서 아래 정보를 정밀 추출하여 오직 JSON 형식으로만 답변하세요.

추출 항목:
1. deposit: 보증금 (예: "40,000,000원 (4,000만원)" 또는 "40,000,000원")
2. monthly_rent: 월 임대료/차임 (예: "3,500,000원 (VAT 별도)" 또는 "3,500,000원")
3. lease_period: 임대기간 (예: "2021-07-31 ~ 2023-07-30 (24개월)")
4. pay_day: 매월 납부일 (예: "매월 31일")
5. landlord: 임대인 성명 (공동명의 포함, 예: "박동석, 김인숙")
6. lessee: 임차인 성명 (예: "주식회사 폭스에듀")
7. bank: 입금 은행 (예: "국민은행" 또는 "확인 필요")
8. account_num: 계좌번호 (예: "123-456-7890" 또는 "확인 필요")
9. account_holder: 예금주 (예: "박동석" 또는 "확인 필요")
10. special_terms: 기타 특약 및 참조사항 (배열 형식, 예: ["1. 계약일 현재 등기부등본 확인 후...", "2. 원상복구..."])

[계약서 텍스트]:
{full_text[:4000]}
"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
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
        print("API Error:", e)
        return None

pdf_p = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\01_최초임대차_판교_판교동612_[박동석-㈜폭스에듀]_(210731).pdf"
parsed = extract_with_gemini_api(pdf_p)
print("\n=== GEMINI LLM PARSED RESULT ===")
print(json.dumps(parsed, ensure_ascii=False, indent=2))
