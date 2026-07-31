# 🌐 AX-Work | Universal Enterprise Asset Management Standard (Master Repo)

**AX-Work**는 특정 법인(기업)에 종속되지 않고, 모든 기업의 자산(부동산, 법인차량, 보험 등)을 효율적으로 관리·파싱·보고서화하기 위한 **독립적인 범용 마스터 표준 저장소**입니다.

---

## 🎯 마스터 저장소 설계 원칙
1. **독립성 (Independence)**: 자산의 개수, 특정 기업명, 특정 법인의 개별 규칙에 종속되지 않으며 범용성을 보장합니다.
2. **표준화 (Standardization)**: 파일명 명명 규칙, 2단계 마스터 폴더 계층 구조, AI OCR 파싱 필드 및 보고서 서식을 일관되게 제공합니다.
3. **선순환 이관 체계 (Continuous Feedback Loop)**: 
   * 실무 전용 저장소(예: FoxConnect-AX)에서 실무를 수행하며 검증된 최적의 프롬프트, 코드, 양식 아이디어를 수집합니다.
   * 기존 마스터 표준과의 비교 분석 후 관리자 승인을 거쳐 본 AX-Work 저장소로 범용화 반영됩니다.

---

## 📁 저장소 아키텍처

`	ext
AX-Work/
├── README.md                                      <-- 마스터 저장소 개요 및 운용 원칙
├── Universal_Enterprise_Asset_Management_Framework.md <-- 범용 기업 자산관리 마스터 프레임워크
├── Real_Estate_Management_Rules.md               <-- 부동산 자산 세부 실무 및 보고서 규칙
├── Antigravity_CLI_GoogleDrive_Setup_Guide.md     <-- Google Drive 연동 및 CLI 구축 가이드
├── main.py                                         <-- 범용 마스터 실행 엔드포인트
├── mcp.json.template                              <-- Google Drive MCP 설정 샘플
├── .gitignore                                      <-- 개인정보 및 보안 파일 업로드 방지 규칙
├── src/                                            <-- 자산 관리 핵심 파이썬 엔진
│   ├── config.py                                   <-- 동적 환경 설정 모듈
│   ├── db_manager.py                               <-- 범용 SQLite DB 엔진
│   ├── folder_structure_engine.py                  <-- 폴더 계층 자동 생성 엔진
│   ├── real_estate_engine.py                       <-- 부동산 파싱/관리 엔진
│   ├── vehicle_engine.py                           <-- 법인 차량 파싱/관리 엔진
│   ├── insurance_engine.py                         <-- 법인 보험 파싱/관리 엔진
│   └── template_exporter.py                        <-- 보고서 양식 내보내기 엔진
└── templates/                                      <-- 마스터 서식 템플릿 저장소
    ├── Master_Corporate_Contract_Note_Overhead_Template.md
    ├── Corporate_Real_Estate_Contract_Note_Template.md
    ├── Corporate_Vehicle_Contract_Note_Template.md
    ├── Corporate_Insurance_Contract_Note_Template.md
    ├── External_Audit_IPO_Asset_Register_Template.md
    └── README.md
`

---

## 🔄 마스터 표준 이관 (Feedback Loop) 프로세스

`	ext
[ FoxConnect-AX 실무 수행 ]
          │
          ▼ (우수 양식/코드/프롬프트 발굴)
[ AI 에이전트 비교 분석 ] ──> 기존 AX-Work 파일과 차이점 & 범용성 비교 보고서 작성
          │
          ▼ (관리자 검토 & 승인)
[ AX-Work 마스터 반영 ] ──> 교체 또는 하단 추가를 통한 마스터 표준 고도화
`

---

## 🛡️ 보안 및 오픈소스 표준 준수
- 실제 기업 계약서 데이터, .csv, .xlsx, .docx, .pdf, .db 등 개별 법인 데이터는 .gitignore에 의해 본 저장소에서 자동 제외됩니다.
- 본 저장소의 모든 변경은 관리자 승인 절차를 거쳐 반영됩니다.
