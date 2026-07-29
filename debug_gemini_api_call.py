import os
import sys
import json
import fitz
import urllib.request
import urllib.error

sys.stdout.reconfigure(encoding='utf-8')

api_key = ""
key_file = r"C:\Users\User\OneDrive\바탕 화면\업무_AX\local_api_key.txt"
if os.path.exists(key_file):
    with open(key_file, "r", encoding="utf-8") as f:
        api_key = f.read().strip()

pdf_path = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_판교_판교동612\01_최초임대차_판교_판교동612_[박동석-㈜폭스에듀]_(210731).pdf"

doc_pdf = fitz.open(pdf_path)
full_text = ""
for page in doc_pdf:
    full_text += page.get_text() + "\n"

prompt = f"""
당신은 대한민국 부동산 임대차계약서 정밀 추출 전문 Vision AI입니다.
아래 계약서 텍스트 원문에서 정보를 추출하여 JSON 형태로 답변해 주세요:

1. deposit: 보증금 금액
2. monthly_rent: 월 임대료/차임
3. lease_period: 임대 기간
4. pay_day: 매월 납부일
5. landlord: 임대인 성명
6. lessee: 임차인 성명
7. bank: 입금 은행
8. account_num: 계좌번호
9. account_holder: 예금주
10. special_terms: 기타 특약사항

[계약서 텍스트]:
{full_text[:3000]}
"""

url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
payload = {
    "contents": [{"parts": [{"text": prompt}]}]
}

req = urllib.request.Request(
    url,
    data=json.dumps(payload).encode('utf-8'),
    headers={'Content-Type': 'application/json'}
)

try:
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode('utf-8'))
        print("Success!")
        print(data['candidates'][0]['content']['parts'][0]['text'])
except urllib.error.HTTPError as e:
    print("HTTPError Code:", e.code)
    print("HTTPError Reason:", e.read().decode('utf-8'))
