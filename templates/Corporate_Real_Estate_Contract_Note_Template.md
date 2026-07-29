# 📄 Universal Corporate Real Estate Asset Management Specification
## 범용 기업 부동산 계약관리 마스터 서식 및 1:1 평면 폴더 규격서

> **공통 최상위 규격**: [`Master_Corporate_Contract_Note_Overhead_Template.md`](Master_Corporate_Contract_Note_Overhead_Template.md) 수칙 100% 준수

---

### 1. 개요 (Overview)
본 표준 명세서는 기업의 부동산(임대차, 전대차, 소유권/매매) 자산을 1:1 독립 평면 폴더 체계로 정돈하고, **`부동산 계약관리 마스터 서식` (Executive Real Estate Management Note)**을 자동으로 생성·유지하기 위한 표준 규격서입니다.

---

### 2. 폴더 명정 및 파일명 표준 (Folder & Filename Standards)

1. **1:1 평면 독립 폴더 구조 (Flat Directory)**:
   - 임대차 폴더: `01_부동산_자산관리_임대차계약\[순번]_[물건명]`
   - 매매/소유권 폴더: `01_부동산_자산관리_매매_소유권문서\[순번]_[물건명]`

2. **Option A 표준 PDF 파일명**:
   - 임대차: `[순번]_[계약유형]_[물건명]_[임대인-임차인]_(YYMMDD).pdf`
   - 전대차: `[순번]_[전대차계약]_[물건명_층수]_[전대인-전차인]_(YYMMDD).pdf`

---

### 3. 부동산 계약관리 마스터 서식 표 구조 (Executive 4-Table Layout)

- **Table 0 (계약 정보)**: 연도, 순서, 구분, 물건명/주소, 임대인, 임대기간, 보증금, 월세
- **Table 1 (수록 계약서 문서 목록)**: 순서, 문서명 (Option A 파일명 연동), 계약 종류, 계약 당사자, 계약 시작일
- **Table 2 (임대인 및 납부 계좌 정보)**: 임대인, 임차인, 연락처, 입금은행, 예금주, 계좌번호, 비고 (관리자 참고사항)
- **Table 3 (계약 변동 이력 및 특이사항)**: 변동/전대차 이력, 연장/갱신 이력, 해지/퇴거 메모, 기타 특약사항
