import os
import sys
import json
import time
import fitz
from google import genai
from google.genai import types
from templates.generate_docx_note import create_contract_note

sys.stdout.reconfigure(encoding='utf-8')

USER_API_KEY = os.environ.get("GEMINI_API_KEY", "")
client = genai.Client(api_key=USER_API_KEY)

FOXCONNECT_ROOT = r"G:\내 드라이브\[FoxConnect]"

PROMPT_VISION = """
당신은 대한민국 부동산 임대차 계약서 정밀 파싱 전문 Vision AI입니다.
첨부된 계약서 스캔본 이미지를 시각적으로 정밀하게 분석하여 아래 JSON 형식에 맞춰 핵심 정보를 정형화 데이터로 추출해 주세요.

[파싱 및 정밀 추론 규칙]
1. 건물명, 호수, 임대인, 임차인, 보증금(원화 숫자만), 월임대료(원화 숫자만), 매월 납부일자(숫자만), 계약면적을 정확히 추출하세요.
2. 파일명 날짜가 아닌, 계약서 본문 제2조(임대차기간) 조항에 명시된 실제 임대 시작일자(lease_start_date)와 임대 종료일자(lease_end_date)를 파싱하세요.
3. 종료일자가 안 적혀있고 '24개월' 또는 '2년'으로만 약정된 경우, 시작일 기준으로 종료일자(YYYY-MM-DD)를 직접 계산하세요.

반환 JSON 형식:
{
  "building_name": "건물명",
  "unit_name": "호수",
  "contract_type": "최초임대차 / 전대차 / 연장계약 등",
  "lessor": "임대인 명칭",
  "lessee": "임차인 명칭",
  "contract_date": "YYYY-MM-DD",
  "lease_start_date": "YYYY-MM-DD",
  "lease_end_date": "YYYY-MM-DD",
  "duration_months": 24,
  "deposit_krw": 46000000,
  "monthly_rent_krw": 4600000,
  "payment_day": 10,
  "area_m2": 218.7,
  "special_notes": "특약사항"
}
"""

def parse_pdf_page_with_vision_safe(pdf_path, max_retries=3):
    img_bytes = None
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=300)
        img_bytes = pix.tobytes("png")
    except Exception as e:
        print(f"  [PDF READ ERROR] {os.path.basename(pdf_path)}: {e}")
        return None

    # Try models in order: gemini-2.0-flash, gemini-flash-latest
    target_models = ['gemini-2.0-flash', 'gemini-flash-latest']

    for model_name in target_models:
        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[
                        types.Part.from_bytes(data=img_bytes, mime_type='image/png'),
                        PROMPT_VISION
                    ]
                )
                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text.replace("```json", "").replace("```", "").strip()
                elif raw_text.startswith("```"):
                    raw_text = raw_text.replace("```", "").strip()
                return json.loads(raw_text)
            except Exception as e:
                err_msg = str(e)
                if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                    print(f"  [RATE LIMIT 429] Retrying {os.path.basename(pdf_path)} with {model_name} in 12s... (Attempt {attempt+1}/{max_retries})")
                    time.sleep(12)
                else:
                    print(f"  [MODEL ERROR {model_name}] {err_msg}")
                    break

    return None

processed_folders = 0
updated_notes = 0

print("Starting Rate-Limit-Safe Vision AI Batch Processing across G:\\내 드라이브\\[FoxConnect]...")

for root, dirs, files in os.walk(FOXCONNECT_ROOT):
    pdf_files = [f for f in files if f.lower().endswith(".pdf")]
    if not pdf_files:
        continue

    pdf_files.sort()
    print(f"\n==========================================")
    print(f"Processing Folder: {root}")
    print(f"Found {len(pdf_files)} PDF contracts.")
    
    parsed_docs = []
    for pdf_f in pdf_files:
        pdf_p = os.path.join(root, pdf_f)
        print(f"  --> Vision AI Parsing: {pdf_f}")
        v_res = parse_pdf_page_with_vision_safe(pdf_p)
        time.sleep(3) # Polite delay between calls
        
        if v_res:
            v_res["raw_filename"] = pdf_f
            parsed_docs.append(v_res)

    if not parsed_docs:
        continue

    primary = parsed_docs[0]
    folder_name = os.path.basename(root)
    parent_name = os.path.basename(os.path.dirname(root))

    dep_val = f"{primary.get('deposit_krw', 0):,} 원" if isinstance(primary.get('deposit_krw'), (int, float)) and primary.get('deposit_krw', 0) > 0 else ""
    rent_val = f"{primary.get('monthly_rent_krw', 0):,} 원" if isinstance(primary.get('monthly_rent_krw'), (int, float)) and primary.get('monthly_rent_krw', 0) > 0 else ""
    pay_day_val = f"매월 {primary.get('payment_day')}일" if primary.get('payment_day') else ""
    
    p_start = primary.get('lease_start_date', primary.get('contract_date', ''))
    p_end = primary.get('lease_end_date', '')
    period_val = f"{p_start} ~ {p_end}" if p_end else f"{p_start} ~ "

    area_val = f"{primary.get('area_m2')} ㎡" if primary.get('area_m2') else ""

    master_info = {
        "building_name": parent_name,
        "unit_name": folder_name,
        "usage": f"{folder_name} 사무실",
        "contract_type": primary.get("contract_type", "최초임대차"),
        "initial_date": primary.get("contract_date", ""),
        "period": period_val,
        "payment_day": pay_day_val,
        "deposit": dep_val,
        "rent": rent_val,
        "area_m2": area_val,
        "area_pyung": "",
        "lessor": primary.get("lessor", ""),
        "lessee": primary.get("lessee", ""),
        "lessor_phone": "",
        "mgmt_phone": "",
        "bank": "",
        "account_holder": primary.get("lessor", ""),
        "account_number": "",
        "remarks": primary.get("special_notes", "특이사항 없음"),
        "history_text": "특이 변동이력 없음 (최초 계약 유지 중)",
        "renewal_text": "",
        "termination_text": "",
        "special_notes": primary.get("special_notes", "")
    }

    docs_list = []
    for p_doc in parsed_docs:
        fn_no_ext = os.path.splitext(p_doc["raw_filename"])[0]
        c_type = p_doc.get("contract_type", "계약서")
        u_str = folder_name
        d_title = f"{u_str} {c_type} 계약서"
        
        l_name = p_doc.get("lessor", "")
        r_name = p_doc.get("lessee", "")
        parties_str = f"{l_name} → {r_name}" if l_name and r_name else l_name

        docs_list.append({
            "display_title": d_title,
            "filename_no_ext": fn_no_ext,
            "contract_type": c_type,
            "parties": parties_str,
            "contract_date": p_doc.get("contract_date", "")
        })

    import re
    clean_u_name = re.sub(r'[\\/\:\*\?\"\<\>\|]', '_', folder_name)
    output_docx = os.path.join(root, f"부동산_계약관리노트_{clean_u_name}.docx")

    existing_notes = [f for f in files if f.lower().endswith(".docx") and ("계약관리노트" in f or "계약_노트" in f or "노트" in f)]
    for old_f in existing_notes:
        try:
            os.remove(os.path.join(root, old_f))
        except Exception:
            pass

    try:
        create_contract_note(master_info, docs_list, output_docx)
        print(f"  ==> [VISION LLM DOCX CREATED] {os.path.basename(output_docx)}")
        updated_notes += 1
    except Exception as e:
        print(f"  ==> [DOCX BUILD ERROR] {e}")

    processed_folders += 1

print("\n==========================================")
print(f"Finished Vision LLM Batch Processing across all {processed_folders} folders!")
print(f"Total updated contract notes created: {updated_notes}")
print("==========================================")
