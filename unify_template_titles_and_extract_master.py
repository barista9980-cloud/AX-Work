import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
templates_dir = os.path.join(REPO_DIR, "templates")

# 1. Create Shared Master Overhead Template
master_overhead_p = os.path.join(templates_dir, "Master_Corporate_Contract_Note_Overhead_Template.md")

master_overhead_content = """# 🏛️ Master Corporate Contract Note Overhead Specification
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
"""

with open(master_overhead_p, "w", encoding="utf-8") as f:
    f.write(master_overhead_content)

print("Created Master_Corporate_Contract_Note_Overhead_Template.md!")

# 2. Update Real Estate Template
re_p = os.path.join(templates_dir, "Corporate_Real_Estate_Contract_Note_Template.md")
re_content = """# 📄 Universal Corporate Real Estate Asset Management Specification
## 범용 기업 부동산 계약관리 마스터 서식 및 1:1 평면 폴더 규격서

> **공통 최상위 규격**: [`Master_Corporate_Contract_Note_Overhead_Template.md`](Master_Corporate_Contract_Note_Overhead_Template.md) 수칙 100% 준수

---

### 1. 개요 (Overview)
본 표준 명세서는 기업의 부동산(임대차, 전대차, 소유권/매매) 자산을 1:1 독립 평면 폴더 체계로 정돈하고, **`부동산 계약관리 마스터 서식` (Executive Real Estate Management Note)**을 자동으로 생성·유지하기 위한 표준 규격서입니다.

---

### 2. 폴더 명정 및 파일명 표준 (Folder & Filename Standards)

1. **1:1 평면 독립 폴더 구조 (Flat Directory)**:
   - 임대차 폴더: `01_부동산_자산관리\01_임대차계약\[순번]_[물건명]`
   - 매매/소유권 폴더: `01_부동산_자산관리\02_매매_소유권문서\[순번]_[물건명]`

2. **Option A 표준 PDF 파일명**:
   - 임대차: `[순번]_[계약유형]_[물건명]_[임대인-임차인]_(YYMMDD).pdf`
   - 전대차: `[순번]_[전대차계약]_[물건명_층수]_[전대인-전차인]_(YYMMDD).pdf`

---

### 3. 부동산 계약관리 마스터 서식 표 구조 (Executive 4-Table Layout)

- **Table 0 (계약 정보)**: 연도, 순서, 구분, 물건명/주소, 임대인, 임대기간, 보증금, 월세
- **Table 1 (수록 계약서 문서 목록)**: 순서, 문서명 (Option A 파일명 연동), 계약 종류, 계약 당사자, 계약 시작일
- **Table 2 (임대인 및 납부 계좌 정보)**: 임대인, 임차인, 연락처, 입금은행, 예금주, 계좌번호, 비고 (관리자 참고사항)
- **Table 3 (계약 변동 이력 및 특이사항)**: 변동/전대차 이력, 연장/갱신 이력, 해지/퇴거 메모, 기타 특약사항
"""

with open(re_p, "w", encoding="utf-8") as f:
    f.write(re_content)

print("Updated Corporate_Real_Estate_Contract_Note_Template.md!")

# 3. Update Vehicle Template
veh_p = os.path.join(templates_dir, "Corporate_Vehicle_Contract_Note_Template.md")
veh_content = """# 📄 Universal Corporate Vehicle Asset Management Specification
## 범용 기업 법인차량 계약관리 마스터 서식 및 1:1 평면 폴더 규격서

> **공통 최상위 규격**: [`Master_Corporate_Contract_Note_Overhead_Template.md`](Master_Corporate_Contract_Note_Overhead_Template.md) 수칙 100% 준수

---

### 1. 개요 (Overview)
본 표준 명세서는 기업의 법인차량(운용리스, 금융리스, 장기렌트 등) 자산을 1:1 독립 평면 폴더 체계로 정돈하고, **`법인차량 계약관리 마스터 서식` (Executive Vehicle Management Note)**을 자동으로 생성·유지하기 위한 표준 규격서입니다.

---

### 2. 폴더 명정 및 파일명 표준 (Folder & Filename Standards)

1. **1:1 평면 독립 폴더 구조 (Flat Directory)**:
   - 최상위 디렉토리: `02_차량_자산관리\01_차량계약_리스_렌트`
   - 차량별 1:1 폴더명: `[순서]_[차종]_[차량번호]([금융사]_[계약유형])`

2. **Option A 표준 PDF 파일명**:
   - 파일명 구조: `[순서]_[차량계약유형]_[차종_차량번호]_[금융사-법인명]_(YYMMDD).pdf`

---

### 3. 법인차량 계약관리 마스터 서식 표 구조 (Executive 4-Table Layout)

- **Table 0 (차량 계약 정보)**: 차종/차량번호, 실사용자, 계약유형, 계약시작일, 약정기간, 보증금/선납금, 월 렌탈/리스료
- **Table 1 (수록 차량 계약서 문서 목록)**: 순서, 문서명 (Option A 파일명 연동), 계약 종류, 계약 당사자, 계약 시작일
- **Table 2 (여신 금융사 및 납부 정보)**: 금융사, 계약법인, 담당자, 자동차보험, 자동이체 은행, 예금주, 월 납입금액, 비고
- **Table 3 (차량 계약 변동 및 만기 이력)**: 승계/양도/양수 이력, 연장/재계약 이력, 중도해지/반납 메모, 기타 약정사항
"""

with open(veh_p, "w", encoding="utf-8") as f:
    f.write(veh_content)

print("Updated Corporate_Vehicle_Contract_Note_Template.md!")

# 4. Update Insurance Template
ins_p = os.path.join(templates_dir, "Corporate_Insurance_Contract_Note_Template.md")
ins_content = """# 📄 Universal Corporate Insurance Asset Management Specification
## 범용 기업 기업보험 계약관리 마스터 서식 및 1:1 평면 폴더 규격서

> **공통 최상위 규격**: [`Master_Corporate_Contract_Note_Overhead_Template.md`](Master_Corporate_Contract_Note_Overhead_Template.md) 수칙 100% 준수

---

### 1. 개요 (Overview)
본 표준 명세서는 기업의 보험 자산(경영인정기보험, 화재/배상책임, 법인차량 자동차보험 등)을 1:1 독립 평면 폴더 체계로 정돈하고, **`기업보험 계약관리 마스터 서식` (Executive Insurance Management Note)**을 자동으로 생성·유지하기 위한 표준 규격서입니다.

---

### 2. 폴더 명정 및 파일명 표준 (Folder & Filename Standards)

1. **1:1 평면 독립 폴더 구조 (Flat Directory)**:
   - 최상위 디렉토리: `03_보험_자산관리\01_보험증권_및_배서계약서`
   - 보험별 1:1 폴더명: `[순서]_[보험종목]_[보험대상/차종]`

2. **Option A 표준 PDF 및 배서 통합 규칙**:
   - 최초 증권 PDF 및 배서 승인서 PDF를 동일한 1:1 보험 폴더 내에 일괄 보관합니다.

---

### 3. 기업보험 계약관리 마스터 서식 표 구조 (Executive 4-Table Layout)

- **Table 0 (보험 계약 정보)**: 보험종목, 피보험자/보장대상, 보험사, 상품명, 증권번호, 보험기간, 납입액/납입주기
- **Table 1 (수록 보험 증권 및 배서 문서 목록)**: 순서, 문서명 (Option A 파일명 연동), 구분, 작성일자
- **Table 2 (보험사 및 납부 계좌 정보)**: 보험사, 담당자, 자동이체 계좌, 예금주, 납입방법, 비고 (세무/보장 참고사항)
- **Table 3 (보험 계약 변동 및 배서 이력)**: 배서/변경 이력, 갱신 기록, 사고/청구 이력, 기타 참조사항
"""

with open(ins_p, "w", encoding="utf-8") as f:
    f.write(ins_content)

print("Updated Corporate_Insurance_Contract_Note_Template.md!")
