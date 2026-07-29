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
  "building_name": "가산 대륭포스트타워6차",
  "unit_name": "402_403호",
  "contract_type": "최초임대차",
  "lessor": "㈜엠씨에스솔루션",
  "lessee": "㈜폭스에듀",
  "lessor_phone": "02-3275-1190",
  "mgmt_phone": "02-3275-1190",
  "bank_name": "신한은행",
  "account_holder": "㈜엠씨에스솔루션",
  "account_number": "415-890001-13104",
  "contract_date": "2024-02-29",
  "lease_start_date": "2024-03-01",
  "lease_end_date": "2026-02-28",
  "deposit_krw": 46000000,
  "monthly_rent_krw": 4600000,
  "payment_day": 10,
  "area_m2": 218.7,
  "area_pyung": 66.15,
  "special_terms": [
    "1. 특약사항 1선",
    "2. 특약사항 2선"
  ]
}
"""

results = []

for pdf_p in pdf_402_403_files:
    print(f"\nParsing: {os.path.basename(pdf_p)}")
    success = False
    for attempt in range(5):
        try:
            doc = fitz.open(pdf_p)
            pix = doc[0].get_pixmap(dpi=300)
            img_bytes = pix.tobytes("png")
            
            response = client.models.generate_content(
                model='gemini-3.5-flash',
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
            print("=== SUCCESS ===")
            print(json.dumps(data, ensure_ascii=False, indent=2))
            success = True
            break
        except Exception as e:
            print(f"Attempt {attempt+1} failed: {e}. Waiting 35 seconds...")
            time.sleep(35)

with open("parsed_402_403_vision.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
