import os
import sys
import json
import fitz
from google import genai
from google.genai import types

sys.stdout.reconfigure(encoding='utf-8')

user_api_key = os.environ.get("GEMINI_API_KEY", "")

print("Testing user API key with model: gemini-3.5-flash ...")

client = genai.Client(api_key=user_api_key)

sample_pdf = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_가산_대륭포스트타워6차\402_403호\가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229).pdf"

doc = fitz.open(sample_pdf)
page = doc[0]
pix = page.get_pixmap(dpi=300)
img_bytes = pix.tobytes("png")

prompt = """
당신은 대한민국 부동산 임대차 계약서 정밀 추출 전문 Vision AI입니다.
첨부된 계약서 스캔본 이미지를 시각적으로 정확히 분석하여 아래 JSON 형식에 맞춰 정형화 데이터를 추출해 주세요.

오직 순수한 JSON 데이터만 반환하세요:
{
  "building_name": "건물명",
  "unit_name": "호수",
  "contract_type": "최초임대차",
  "lessor": "임대인",
  "lessee": "임차인",
  "contract_date": "YYYY-MM-DD",
  "lease_start_date": "YYYY-MM-DD",
  "lease_end_date": "YYYY-MM-DD",
  "duration_months": 24,
  "deposit_krw": 46000000,
  "monthly_rent_krw": 4600000,
  "payment_day": 10
}
"""

try:
    response = client.models.generate_content(
        model='gemini-3.5-flash',
        contents=[
            types.Part.from_bytes(data=img_bytes, mime_type='image/png'),
            prompt
        ]
    )
    
    print("\n=== GEMINI 3.5 FLASH VISION RESPONSE SUCCESS ===")
    print(response.text)
    
except Exception as e:
    print("Error calling Gemini 3.5 Flash:", str(e))
