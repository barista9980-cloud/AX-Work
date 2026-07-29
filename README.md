# 🏢 FoxConnect 부동산 자산관리 AX 자동화 파이프라인 & 마스터 템플릿

본 저장소(Repository)는 **FoxConnect 부동산 자산관리 업무의 AX(AI Transformation) 전환**을 위한 **계약관리노트 표준 마스터 서식**과 **Vision LLM 기반 자동화 파이프라인 엔진**의 전체 규격 및 소스 코드를 관리합니다.

---

## 📌 주요 업데이트 내역 (Today's Updates)

### 1. 부동산 계약관리노트 마스터 양식 표준화 (Executive Master Template)
- **표 전체 가로 너비 통일 (17.4 cm / 6.85 inch)**:
  - 1~4번 모든 표의 전체 너비를 **17.4 cm로 1.0mm 오차 없이 100% 동일 일치**시켜 좌·우측 테두리선 수직 칼정렬 완성.
- **워드 좌측 검은 점 기호(`▪`) 영구 전면 제거**:
  - `keep_with_next = False` 및 XML 노드 전면 삭제를 통해 편집 시 지워지지 않던 좌측 비인쇄 검은 점 기호를 100% 차단.
- **클린 2페이지 모듈 분할 (Clean 2-Page Structure)**:
  - 표가 페이지 경계에서 어중간하게 잘리는 현상을 원천 차단하기 위해 섹션 3(`3. 임대인 및 납부 계좌 정보`) 시작 전 **명시적 Page Break를 적용**.
  - **Page 1**: 대제목 + `1. 계약 정보` (Table 0) + `2. 수록 계약서 문서 목록` (Table 1)
  - **Page 2**: `3. 임대인 및 납부 계좌 정보` (Table 2) + `4. 계약 변동 이력 및 특이사항` (Table 3)
- **타이포그래피 & 수동 입력 최적화**:
  - 대제목 `13pt Bold`, 소제목 `11pt Bold` (`1.`, `2.`, `3.`, `4.` 마침표 표기), 셀 본문 및 빈 셀 `10pt 맑은 고딕` 통일 (관리자 수동 작성 최적화).
- **수록 계약서 문서 목록 시각적 개편**:
  - 대표 문서명(Bold 10pt) + 원천 파일명(Gray 8.5pt) **2단 구성 및 얼룩말 패턴 교차 배색** 적용.

---

### 2. Vision LLM 기반 정밀 파싱 & AX 파이프라인 초기 셋팅 (Vision LLM AX Pipeline)
- **단순 OCR의 한계 극복**:
  - 스캔 이미지 PDF의 글자 잘림, 오타, 표 엉킴 문제를 멀티모달 비전 AI(Gemini Vision API)를 통해 시각적으로 직접 인지하여 정밀 추출.
- **임대차 기간 암묵적 수학적 계산 (Implicit Reasoning)**:
  - 계약서 본문에 `"시작일 2024-03-01, 약정기간 24개월"`만 기재되어 있어도 Vision AI가 스스로 **종료일자(`2026-02-28`)를 자동 산출하여 정형 JSON(`lease_end_date`)으로 반환**.
- **안전한 Rate-Limit (429) 백오프 핸들링**:
  - API 호출 간격 조율(`time.sleep`) 및 자동 재시도(Exponential Backoff) 로직 내장.

---

## 📐 마스터 템플릿 명세서 (Template Specification)

상세 마스터 템플릿 규격은 [`templates/Real_Estate_Contract_Note_Template.md`](templates/Real_Estate_Contract_Note_Template.md) 문서에서 확인하실 수 있습니다.

---

## 🛠️ 주요 파이썬 엔진 모듈 (Python Modules)

- **`templates/generate_docx_note.py`**: 마스터 서식 및 2페이지 분할 규칙이 적용된 `.docx` 자동 생성 모듈
- **`vision_llm_contract_parser.py`**: PDF 스캔본 300DPI 렌더링 & Vision LLM API 페이로드 생성 모듈
- **`batch_process_with_rate_limit.py`**: 전체 세부 물건 구글 드라이브 일괄 Vision AI 파싱 및 문서 자동 갱신 엔진
