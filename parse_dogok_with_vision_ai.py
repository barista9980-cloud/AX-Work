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

target_dir = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\02_강남_도곡로1길23\지하1층,1층,2층,3층"

pdf_files = [
    "강남_도곡로1길23_전층_01_최초임대차_[박재윤-㈜폭스에듀]_(241107).pdf",
    "강남_도곡로1길23_전층_02_변경계약_[유한회사 청송(박재윤)-㈜폭스에듀]_(250901).pdf",
    "강남_도곡로1길23_1층_01_전대차_[㈜폭스에듀-㈜에스앤에이치트레이딩]_(241101).pdf",
    "강남_도곡로1길23_1층_02_전대차_[㈜폭스에듀-한국경찰과학전략센터]_(250821).pdf",
    "강남_도곡로1길23_1층_03_전대차_[㈜폭스에듀-㈜월드유니코어]_(250821).pdf",
    "강남_도곡로1길23_2층_01_전대차_[㈜폭스에듀-㈜실리콘아츠]_(241101).pdf",
    "강남_도곡로1길23_2층_02_전대차_[㈜폭스에듀-㈜하이퍼비주얼에이아이]_(250101).pdf",
    "강남_도곡로1길23_3층_01_전대차_[㈜폭스에듀-㈜트라이디스]_(250124).pdf"
]

prompt = """
대한민국 부동산 임대차/전대차 계약서 스캔본 이미지에서 다음 핵심 정보를 정밀 추출하세요.

1. 건물명/호수, 계약유형 (최초임대차/전대차/변경계약)
2. 임대인 (전대인), 임차인 (전차인), 임대인연락처, 관리사무소연락처, 입금은행, 예금주, 계좌번호
3. 보증금 (원), 월 임대료 (원), 매월 납부일자 (일)
4. 계약 체결일자 (YYYY-MM-DD), 실제 임대 시작일자 (YYYY-MM-DD), 실제 임대 종료일자 (YYYY-MM-DD, 약정기간 24개월 시 계산)
5. 본문 주요 특약사항 (1. 2. 3. 번호목록 형태)

오직 순수한 JSON만 반환하세요:
{
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
  "payment_day": "",
  "special_terms": [
    "1. 특약사항 1",
    "2. 특약사항 2"
  ]
}
"""

parsed_results = {}

for pdf_name in pdf_files:
    pdf_path = os.path.join(target_dir, pdf_name)
    print(f"\n==================================================")
    print(f"Vision AI Parsing: {pdf_name}")
    print(f"==================================================")
    
    if not os.path.exists(pdf_path):
        print("File missing:", pdf_path)
        continue

    for attempt in range(5):
        try:
            doc = fitz.open(pdf_path)
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
            parsed_results[pdf_name] = data
            print("=== VISION AI SUCCESS ===")
            print(json.dumps(data, ensure_ascii=False, indent=2))
            break
        except Exception as e:
            print(f"Attempt {attempt+1} failed: {e}. Waiting 30s...")
            time.sleep(30)

with open("dogok_vision_ai_parsed_results.json", "w", encoding="utf-8") as f:
    json.dump(parsed_results, f, ensure_ascii=False, indent=2)

print("\nParsing complete! Saved to dogok_vision_ai_parsed_results.json")
