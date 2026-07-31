# 🏛️ Master Corporate Contract Note Overhead Specification
> **범용 기업 계약관리노트 최상위 표준 마스터 명세서**  
> **Applicable Domains**: 부동산 (Real Estate), 법인차량 (Vehicles), 기업보험 (Insurance)  
> **Reference Benchmark Document**: `01 판교 판교동 612` 계약관리노트 (`부동산_계약관리노트_판교_판교동612.docx`)

---

## 📐 1. 문서 공통 규격 및 타이포그래피 (Master Document Typography)

* **페이지 여백 (Page Margins)**:
  - 위쪽 (Top): **42.5 pt** (15 mm)
  - 아래쪽 (Bottom): **42.5 pt** (15 mm)
  - 왼쪽 (Left): **51.0 pt** (18 mm)
  - 오른쪽 (Right): **51.0 pt** (18 mm)
* **글꼴 규격 (Default Font)**: `맑은 고딕` (Malgun Gothic)
* **대제목 (Document Title)**: **13.0 pt** Bold, `#0F172A` (Dark Slate)
* **AI 배너 문구 (AI Subtitle Banner)**: **9.0 pt** Regular, `#475569` (Muted Slate)
* **섹션 소제목 (Section Headings)**: **11.0 pt** Bold, `#1E293B` (Slate Dark, 마침표 및 숫자인덱스 필수 표기)
* **표 헤더 배경색 (Table Header Shading)**: `#F1F5F9` (Light Slate Gray)
* **표 가로 너비 통일 (Table Width Unification)**:
  - 1~4번 모든 표의 **전체 가로 너비를 493.1 pt (6.85인치 / 17.4 cm)로 수직 칼정렬 중앙배치 (`<w:jc w:val="center"/>`)**
* **표 테두리 (Table Borders)**:
  - 선명한 실선 테두리 (`<w:tblBorders>` color `#334155`, sz `4`)
* **클린 2페이지 모듈 분할 원칙 (Clean 2-Page Structure)**:
  - 섹션 3(`3. 임대인 및 납부 계좌 정보`) 시작 전 **명시적 Page Break 적용**
  - **Page 1**: 대제목 + AI 안내 배너 + `1. 계약 정보` (Table 1: 5행 4열) + `2. 수록 계약서 문서 목록 (총 N건)` (Table 2: N행 5열)
  - **Page 2**: `3. 임대인 및 납부 계좌 정보` (Table 3: 6행 4열) + `4. 계약 변동 이력 및 특이사항` (Table 4: 5행 2열)

---

## 📋 2. 상세 표열 너비 및 타이포그래피 (Exact Table Column Specs)

#### **1. 계약 정보 (Table 1: 5행 4열 - 전체너비 493.1pt)**
* Col 1 (라벨): **100.75 pt** (1.40 in), `#F1F5F9` 배경, 10.0pt Bold `#0F172A`
* Col 2 (값): **145.8 pt** (2.025 in), 흰색 배경, 10.0pt Regular `#0F172A`
* Col 3 (라벨): **100.75 pt** (1.40 in), `#F1F5F9` 배경, 10.0pt Bold `#0F172A`
* Col 4 (값): **145.8 pt** (2.025 in), 흰색 배경, 10.0pt Regular `#0F172A`

#### **2. 수록 계약서 문서 목록 (Table 2: N행 5열 - 전체너비 493.2pt)**
* Col 1 (순서): **32.4 pt** (0.45 in), 중앙 정렬, 10.0pt Regular `#0F172A`
* Col 2 (문서명/파일명 - 2단 타이포그래피): **212.4 pt** (2.95 in)
  - 1번째 줄 (계약서 명칭): **9.5 pt** Bold, `#0F172A`
  - 2번째 줄 (스캔본 파일명): **8.5 pt** Regular, `#64748B`
* Col 3 (계약종류): **61.2 pt** (0.85 in), 중앙 정렬, 10.0pt Regular `#0F172A`
* Col 4 (계약당사자): **126.0 pt** (1.75 in), 좌측 정렬, 10.0pt Regular `#0F172A`
* Col 5 (임대시작일): **61.2 pt** (0.85 in), 중앙 정렬, 10.0pt Regular `#0F172A`

#### **3. 임대인 및 납부 계좌 정보 (Table 3: 6행 4열 - 5,6행 4열 전면 병합)**
* R1~R4: Col 1/3 (100.75pt `#F1F5F9` Bold), Col 2/4 (145.8pt 흰색 Regular)
* R5 (통합 비고 헤더): **493.1 pt** (4개 열 전면 병합), `#F1F5F9` 배경, **10.0pt Bold** `#0F172A` ("비고 (관리자 참고사항)")
* R6 (통합 비고 데이터): **493.1 pt** (4개 열 전면 병합), 흰색 배경, **10.0pt Regular** `#0F172A` (`1.`, `2.`, `3.` 번호 목록)

#### **4. 계약 변동 이력 및 특이사항 (Table 4: 5행 2열 - 4,5행 2열 전면 병합)**
* R1~R3: Col 1 (144.0pt `#F1F5F9` Bold), Col 2 (349.15pt 흰색 Regular)
* R4 (통합 특약 헤더): **493.15 pt** (2개 열 전면 병합), `#F1F5F9` 배경, **10.0pt Bold** `#0F172A` ("기타 특약 및 참조사항")
* R5 (통합 특약 데이터): **493.15 pt** (2개 열 전면 병합), 흰색 배경, **10.0pt Regular** `#0F172A` (`1.`, `2.`, `3.` 번호 목록)

---

## 🤖 3. Vision AI 파싱 및 하이브리드 입력 원칙

1. **Vision AI 100% 자동 파싱**: 계약서/약정서 PDF 원본에서 파싱된 정보는 100% 자동 기재 (`🤖` 기재).
2. **`[관리자 작성 필요]` 표기 최소화**: 오직 화질 저하 또는 원본상 미존재로 AI 판독 불가 시에만 `[관리자 작성 필요]` 표기.
