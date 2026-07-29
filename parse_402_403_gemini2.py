import os
import sys
import fitz
from google import genai
from google.genai import types
import json
import time

sys.stdout.reconfigure(encoding='utf-8')

USER_API_KEY = os.environ.get("GEMINI_API_KEY", "")
client = genai.Client(api_key=USER_API_KEY)

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

pdf_402_403_files = []
for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    if "402_403" in root:
        for f in files:
            if f.endswith(".pdf"):
                pdf_402_403_files.append(os.path.join(root, f))

pdf_402_403_files.sort()

prompt = """
대한민국 부동산 임대차/전대차 계약서 스캔본 이미지에서 모든 항목을 정밀 추출해 주세요.
1. 건물명/호수, 계약유형 (최초임대차/전대차), 임대인(전대인), 임차인(전차인), 임대인연락처, 관리사무소연락처, 입금은행, 예금주, 계좌번호
2. 보증금, 월임대료, 매월 납부일자, 계약면적, 전용면적
3. 계약 작성일자, 실제 임대시작일자, 임대종료일자 (24개월 약정이면 계산)
4. 본문 특약사항 전체 (1. 2. 3. 번호를 붙여서 각 조항별로 나열)

오직 순수한 JSON만 반환하세요:
{
  "building_name": "",
  "unit_name": "",
  "contract_type": "",
  "lessor": "",
  "lessee": "",
  "lessor_phone": "",
  "mgmt_phone": "",
  "bank_name": "",
  "account_holder": "",
  "account_number": "",
  "contract_date": "YYYY-MM-DD",
  "lease_start_date": "YYYY-MM-DD",
  "lease_end_date": "YYYY-MM-DD",
  "deposit_krw": 0,
  "monthly_rent_krw": 0,
  "payment_day": 0,
  "area_m2": 0.0,
  "area_pyung": 0.0,
  "special_terms": [
    "1. 특약사항1",
    "2. 특약사항2"
  ]
}
"""

results = []

for pdf_p in pdf_402_403_files:
    print(f"\nParsing: {os.path.basename(pdf_p)}")
    try:
        doc = fitz.open(pdf_p)
        pix = doc[0].get_pixmap(dpi=300)
        img_bytes = pix.tobytes("png")
        
        response = client.models.generate_content(
            model='gemini-2.0-flash',
            contents=[
                types.Part.from_bytes(data=img_bytes, mime_type='image/png'),
                prompt
            ]
        )
        
        raw_text = response.text.strip()
        if raw_text.startswith("```json"):
            raw_text = raw_text.replace("```json", "").replace("```", "").strip()
        elif raw_text.startswith("```"):
            raw_text = raw_text.replace("```", "").strip()
            
        data = json.loads(raw_text)
        data["filename"] = os.path.basename(pdf_p)
        results.append(data)
        print(json.dumps(data, ensure_ascii=False, indent=2))
        time.sleep(3)
    except Exception as e:
        print("Error:", e)

with open("parsed_402_403_vision.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
