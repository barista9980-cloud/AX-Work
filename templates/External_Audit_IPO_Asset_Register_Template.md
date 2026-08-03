# External Audit & IPO Submission View Specification
## 외부감사(외감) 및 IPO 제출용 뷰 추출 표준 규격서 (Advanced Compliance Version)

---

### 1. 개요 (Overview)
본 명세서는 기업 통합 총괄자산대장(Master Asset Register)을 기반으로, 회계법인의 **외부감사인(External Auditor) 및 상장주관사/한국거래소 IPO 심사관 요청 시 일시적으로 생성/추출하여 제출하는 재무 및 내부통제 검증용 뷰(On-Demand Audit/IPO View)** 규격서입니다.

계약서 PDF 원본 및 워드 계약관리노트(.docx)에서 추출 가능한 데이터 중 **회계감사 및 IPO 상장심사에 결정적인 4대 고급 검증 수식어**를 포함합니다.

---

### 2. 계약서 PDF & 계약노트 추출 4대 핵심 회계/IPO 검증 데이터

#### A. 📑 K-IFRS 1116 리스회계 검증 데이터 (Rent-Free & Lease Options)
- **렌트프리(Rent-Free) 무상임차 기간**: 최초 1~2개월 무상 임대 조건 (사용권자산/리스부채 평준화 인식에 필수).
- **만기 처리 옵션(Extension/Termination Option)**: 만기 시 `인수 / 반납 / 연장` 선택권 및 중도해지 위약금율.

#### B. 👤 특수관계인 거래 지정 여부 (Related Party Transaction Flag - K-IFRS 1024)
- **대주주/임원 관련자 거래 여부**: 임대인, 피보험자, 주 운행자가 법인의 대주주, 대표이사 또는 특수관계인/계열사인지 자동 플래그(`Y/N`).
- **IPO 심사 중요성**: 한국거래소 IPO 상장심사 시 대주주 사익편취 및 부당 지원 거래 검증 1순위 항목.

#### C. 🔒 우발채무 및 담보/보증 설정 여부 (Pledge & Contingent Liabilities)
- **보증금 채권 질권설정(Pledge)**: 임차/리스보증금이 금융기관 대출 담보로 질권 설정되어 있는지 유무.
- **연대보증(Guarantee)**: 대표이사 개인 연대보증 또는 제3자 지급보증 개입 여부.

#### D. 🏗️ 원상복구 충당부채 의무 (Asset Retirement Obligation - ARO)
- **원상복구 의무 조항**: 부동산 임대차 종료 시 시설 철거 및 원상복구 의무 유무 (`원상복구 의무 명시`).

---

### 3. 외부감사 / IPO 제출 뷰의 핵심 필터링 지침

1. **내부 실무 전용 컬럼 숨기기 (Hide Operational Columns)**:
   - 외부 제출 시 개인정보 및 총무 실무 전용 정보인 `주 운행자 / 실사용 부서`, `매월 납부일`, `은 행`, `계 좌 번 호` 컬럼을 **숨기기(Hide)** 처리하여 제출합니다.
2. **재무 및 감사 필터링 뷰 강조 (Financial & Compliance View)**:
   - 시트 1 상단 **Executive Financial Summary (Rows 4~6)**의 기말 재무 잔액 스냅샷 강조.
   - 시트 2 **`📊 회계/재무 검증용 3개년 현금흐름 및 기말 잔액 대사표 (Reconciliation Table)`**를 외부감사인 제출용 메인 증적으로 활용.
3. **추출 파일명 규격 (Export Naming Rule)**:
   - `[외감_IPO제출용]_2026년도_부동산_총괄자산대장_[㈜대상기업].xlsx`
   - `[외감_IPO제출용]_2026년도_법인차량_총괄자산대장_[㈜대상기업].xlsx`
   - `[외감_IPO제출용]_2026년도_기업보험_총괄자산대장_[㈜대상기업].xlsx`
