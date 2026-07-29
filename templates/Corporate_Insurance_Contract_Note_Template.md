# 📄 Universal Corporate Insurance Asset Management Specification
## 범용 기업 기업보험 계약관리 마스터 서식 및 1:1 평면 폴더 규격서

> **공통 최상위 규격**: [`Master_Corporate_Contract_Note_Overhead_Template.md`](Master_Corporate_Contract_Note_Overhead_Template.md) 수칙 100% 준수

---

### 1. 개요 (Overview)
본 표준 명세서는 기업의 보험 자산(경영인정기보험, 화재/배상책임, 법인차량 자동차보험 등)을 1:1 독립 평면 폴더 체계로 정돈하고, **`기업보험 계약관리 마스터 서식` (Executive Insurance Management Note)**을 자동으로 생성·유지하기 위한 표준 규격서입니다.

---

### 2. 폴더 명정 및 파일명 표준 (Folder & Filename Standards)

1. **1:1 평면 독립 폴더 구조 (Flat Directory)**:
   - 최상위 디렉토리: `03_보험_자산관리_보험증권_및_배서계약서`
   - 보험별 1:1 폴더명: `[순서]_[보험종목]_[보험대상/차종]`

2. **Option A 표준 PDF 및 배서 통합 규칙**:
   - 최초 증권 PDF 및 배서 승인서 PDF를 동일한 1:1 보험 폴더 내에 일괄 보관합니다.

---

### 3. 기업보험 계약관리 마스터 서식 표 구조 (Executive 4-Table Layout)

- **Table 0 (보험 계약 정보)**: 보험종목, 피보험자/보장대상, 보험사, 상품명, 증권번호, 보험기간, 납입액/납입주기
- **Table 1 (수록 보험 증권 및 배서 문서 목록)**: 순서, 문서명 (Option A 파일명 연동), 구분, 작성일자
- **Table 2 (보험사 및 납부 계좌 정보)**: 보험사, 담당자, 자동이체 계좌, 예금주, 납입방법, 비고 (세무/보장 참고사항)
- **Table 3 (보험 계약 변동 및 배서 이력)**: 배서/변경 이력, 갱신 기록, 사고/청구 이력, 기타 참조사항
