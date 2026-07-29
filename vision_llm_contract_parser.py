import os
import sys
import json
import fitz # PyMuPDF
import base64

# Vision LLM Contract Parser Blueprint Engine
# Designed for Antigravity AGY / MCP Integration

def convert_pdf_page_to_base64_png(pdf_path, page_num=0, dpi=300):
    """
    Converts a specific page of a PDF document into a high-res PNG base64 string.
    """
    doc = fitz.open(pdf_path)
    if page_num >= len(doc):
        page_num = 0
    page = doc[page_num]
    pix = page.get_pixmap(dpi=dpi)
    img_bytes = pix.tobytes("png")
    b64_str = base64.b64encode(img_bytes).decode('utf-8')
    return b64_str

VISION_LLM_PROMPT_TEMPLATE = """
당신은 대한민국 부동산 임대차 계약서 정밀 추출 전문 Vision AI입니다.
첨부된 계약서 스캔본 이미지를 시각적으로 정확히 분석하여 아래 JSON 형식에 맞춰 정형화 데이터를 추출해 주세요.

[파싱 및 정밀 추론 규칙]
1. 건물명/호수, 임대인, 임차인, 보증금, 월임대료, 매월 납부일자, 계약면적을 정확히 읽어내어 기재하세요.
2. 파일명 끝의 체결일자가 아닌, 계약서 본문 제2조(임대차기간) 조항에 명시된 실제 임대 시작일(lease_start_date)과 임대 종료일(lease_end_date)을 추출하세요.
3. 만약 종료일자가 직접 날짜로 표시되지 않고 '계약일로부터 24개월' 또는 '2년'으로 약정되어 있다면, 시작일 기준으로 정확한 종료일자(YYYY-MM-DD)를 직접 수학적으로 계산하여 기재하세요.
4. 확실하지 않은 정보가 있다면 '확인필요' 문구를 남기고 reasoning_notes에 사유를 기술하세요.

[반환 JSON 스키마]
{
  "building_name": "건물명",
  "unit_name": "호수",
  "contract_type": "최초임대차 / 전대차 / 연장계약 등",
  "lessor": "임대인 명칭",
  "lessee": "임차인 명칭",
  "contract_date": "계약 작성일자 (YYYY-MM-DD)",
  "lease_start_date": "임대 시작일자 (YYYY-MM-DD)",
  "lease_end_date": "임대 종료일자 (YYYY-MM-DD)",
  "duration_months": 24,
  "deposit_krw": 46000000,
  "monthly_rent_krw": 4600000,
  "payment_day": 10,
  "area_m2": 218.7,
  "special_notes": "특약사항 요약",
  "reasoning_notes": "Vision AI 분석 및 날짜 계산 소견"
}
"""

def generate_vision_llm_request_payload(pdf_path):
    """
    Generates a structured payload ready for Vision LLM API (Gemini / GPT-4o / Claude).
    """
    b64_img = convert_pdf_page_to_base64_png(pdf_path, page_num=0, dpi=300)
    
    payload = {
        "prompt": VISION_LLM_PROMPT_TEMPLATE,
        "image_data": {
            "mime_type": "image/png",
            "base64": b64_img
        }
    }
    return payload

if __name__ == "__main__":
    sample_pdf = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_가산_대륭포스트타워6차\402_403호\가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229).pdf"
    if os.path.exists(sample_pdf):
        payload = generate_vision_llm_request_payload(sample_pdf)
        print("=== VISION LLM REQUEST PAYLOAD READY ===")
        print(f"Base64 Image Size: {len(payload['image_data']['base64'])} bytes")
        print("Prompt Sample:\n", payload['prompt'][:250], "...")
