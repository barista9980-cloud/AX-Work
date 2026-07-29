# 📄 Universal Corporate Vehicle Asset Management Specification
## 범용 기업 법인차량 계약관리 마스터 서식 및 1:1 평면 폴더 규격서

> **공통 최상위 규격**: [`Master_Corporate_Contract_Note_Overhead_Template.md`](Master_Corporate_Contract_Note_Overhead_Template.md) 수칙 100% 준수

---

### 1. 개요 (Overview)
본 표준 명세서는 기업의 법인차량(운용리스, 금융리스, 장기렌트 등) 자산을 1:1 독립 평면 폴더 체계로 정돈하고, **`법인차량 계약관리 마스터 서식` (Executive Vehicle Management Note)**을 자동으로 생성·유지하기 위한 표준 규격서입니다.

---

### 2. 폴더 명정 및 파일명 표준 (Folder & Filename Standards)

1. **1:1 평면 독립 폴더 구조 (Flat Directory)**:
   - 최상위 디렉토리: `02_차량_자산관리_차량계약_리스_렌트`
   - 차량별 1:1 폴더명: `[순서]_[차종]_[차량번호]([금융사]_[계약유형])`

2. **Option A 표준 PDF 파일명**:
   - 파일명 구조: `[순서]_[차량계약유형]_[차종_차량번호]_[금융사-법인명]_(YYMMDD).pdf`

---

### 3. 법인차량 계약관리 마스터 서식 표 구조 (Executive 4-Table Layout)

- **Table 0 (차량 계약 정보)**: 차종/차량번호, 실사용자, 계약유형, 계약시작일, 약정기간, 보증금/선납금, 월 렌탈/리스료
- **Table 1 (수록 차량 계약서 문서 목록)**: 순서, 문서명 (Option A 파일명 연동), 계약 종류, 계약 당사자, 계약 시작일
- **Table 2 (여신 금융사 및 납부 정보)**: 금융사, 계약법인, 담당자, 자동차보험, 자동이체 은행, 예금주, 월 납입금액, 비고
- **Table 3 (차량 계약 변동 및 만기 이력)**: 승계/양도/양수 이력, 연장/재계약 이력, 중도해지/반납 메모, 기타 약정사항
