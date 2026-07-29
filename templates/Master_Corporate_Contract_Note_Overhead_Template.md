# 🏛️ Master Corporate Contract Note Overhead Specification
> **범용 기업 계약관리노트 최상위 표준 마스터 명세서**  
> **Applicable Domains**: 부동산 (Real Estate), 법인차량 (Vehicles), 기업보험 (Insurance)

---

## 📐 1. 문서 공통 규격 및 타이포그래피 (Master Document Typography)

* **기본 글꼴 (Default Font)**: `맑은 고딕` (Malgun Gothic) **10.0 pt**
* **대제목 (Document Title)**: **13.0 pt** (Bold, `#0F172A` Slate Dark)
* **소제목 (Section Headings)**: **11.0 pt** (Bold, `#1E293B` Slate, 마침표 `.` 필수 표기)
* **표 내부 텍스트 기본 크기**: **10.0 pt** (수동 작성 시 10pt 맑은 고딕 유지)
* **표 가로 너비 통일 (Table Width Unification)**:
  - 1~4번 모든 표의 **전체 가로 너비를 6.85인치 (17.4 cm)로 100% 수직 칼정렬**
* **표 테두리 (Table Borders)**:
  - 선명한 실선 테두리 (`<w:tblBorders>` color `#334155`) 지정
* **클린 2페이지 모듈 분할 원칙 (Clean 2-Page Structure)**:
  - 섹션 3(`3. 담당자 및 납부 정보`) 시작 전 **명시적 Page Break 적용**
  - **Page 1**: 대제목 + `1. 계약 정보` (Table 0) + `2. 수록 계약서 문서 목록` (Table 1)
  - **Page 2**: `3. 담당자 및 납부 계좌 정보` (Table 2) + `4. 계약 변동 이력 및 특이사항` (Table 3)
* **비고 및 특약 행 확장 규칙 (Enlarged Row & Numbered List)**:
  - **3번 표 비고**: 4개 열 전체를 통합한 전면 한 줄 헤더 + 전면 한 줄 데이터 행 구성 (`1.`, `2.`, `3.` 번호 목록 나열)
  - **4번 표 기타 특약 및 참조사항**: 2개 열 전체를 통합한 전면 한 줄 헤더 + 전면 한 줄 데이터 행 구성 (`1.`, `2.`, `3.` 번호 목록 나열)

---

## 🤖 2. Vision AI 파싱 및 하이브리드 입력 원칙

1. **Vision AI 100% 자동 파싱**: 계약서/약정서 PDF 원본에서 파싱된 정보는 100% 자동 기재 (`🤖` 기재).
2. **`[관리자 작성 필요]` 표기 최소화**: 오직 화질 저하 또는 원본상 미존재로 AI 판독 불가 시에만 `[관리자 작성 필요]` 표기.
