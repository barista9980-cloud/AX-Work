import os
import sys
import json
import fitz # PyMuPDF
import base64

# ==============================================================================
# Vision LLM (Gemini 1.5) 기반 부동산 계약서 정밀 파싱 실행 모듈
# ==============================================================================

try:
    from google import genai
    from google.genai import types
    HAS_GEMINI_SDK = True
except ImportError:
    HAS_GEMINI_SDK = False

def convert_pdf_page_to_png_bytes(pdf_path, page_num=0, dpi=300):
    doc = fitz.open(pdf_path)
    if page_num >= len(doc):
        page_num = 0
    page = doc[page_num]
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png")

PROMPT_VISION_CONTRACT = """
당신은 대한민국 부동산 임대차 계약서 정밀 추출 전문 Vision AI입니다.
첨부된 계약서 스캔본 이미지를 시각적으로 정확히 분석하여 아래 JSON 형식에 맞춰 정형화 데이터를 추출해 주세요.

[파싱 및 정밀 추론 규칙]
1. 건물명/호수, 임대인, 임차인, 보증금, 월임대료, 매월 납부일자, 계약면적을 정확히 읽어내어 기재하세요.
2. 파일명 끝의 체결일자가 아닌, 계약서 본문 제2조(임대차기간) 조항에 명시된 실제 임대 시작일(lease_start_date)과 임대 종료일(lease_end_date)을 추출하세요.
3. 만약 종료일자가 직접 날짜로 표시되지 않고 '계약일로부터 24개월' 또는 '2년'으로 약정되어 있다면, 시작일 기준으로 정확한 종료일자(YYYY-MM-DD)를 직접 수학적으로 계산하여 기재하세요.
4. 확실하지 않은 정보가 있다면 '확인필요' 문구를 남기고 reasoning_notes에 사유를 기술하세요.

응답은 오직 파싱된 순수 JSON 데이터만 반환하세요 (markdown code block ```json ... ``` 포함 가능).
"""

def parse_contract_with_vision_llm(pdf_path, api_key=None):
    if not HAS_GEMINI_SDK:
        return {"error": "google-genai 패키지가 설치되지 않았습니다. 'pip install google-genai' 명령어를 실행해 주세요."}
    
    if not api_key:
        api_key = os.environ.get("GEMINI_API_KEY")
        
    if not api_key:
        return {"error": "GEMINI_API_KEY 환경변수 또는 API 키가 입력되지 않았습니다."}
        
    img_bytes = convert_pdf_page_to_png_bytes(pdf_path, page_num=0, dpi=300)
    
    client = genai.Client(api_key=api_key)
    
    try:
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=[
                types.Part.from_bytes(data=img_bytes, mime_type='image/png'),
                PROMPT_VISION_CONTRACT
            ]
        )
        
        raw_text = response.text.strip()
        # Clean markdown formatting if present
        if raw_text.startswith("```json"):
            raw_text = raw_text.replace("```json", "").replace("```", "").strip()
        elif raw_text.startswith("```"):
            raw_text = raw_text.replace("```", "").strip()
            
        json_data = json.loads(raw_text)
        return json_data
    except Exception as e:
        return {"error": f"Vision LLM API 호출 실패: {str(e)}"}

if __name__ == "__main__":
    print("=== Vision LLM Gemini 파싱 테스트 모듈 ===")
    sample_pdf = r"G:\내 드라이브\[FoxConnect]\[총무]업무\01_부동산_자산관리\01_임대차계약\01_가산_대륭포스트타워6차\402_403호\가산_대륭포스트타워6차_402_403호_01_최초임대차_[㈜엠씨에스솔루션-㈜폭스에듀]_(240229).pdf"
    print("Target PDF:", sample_pdf)
    
    api_key = os.environ.get("GEMINI_API_KEY")
    if api_key:
        print("API Key Detected! Running Vision LLM Parsing...")
        result = parse_contract_with_vision_llm(sample_pdf, api_key)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("GEMINI_API_KEY 환경변수가 아직 설정되지 않았습니다. 가이드 1~3단계를 진행해 주세요.")
