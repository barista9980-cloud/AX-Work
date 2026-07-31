# 🌐 AX-Work | Universal Enterprise Asset Management Standard (Master Repo)

> **독립적 마스터 표준 저장소 (Universal Master Standard Repository)**
> 특정 법인(기업)에 종속되지 않고, 모든 기업의 자산(부동산, 법인차량, 기업보험 등)을 효율적으로 관리·자동화 파싱·보고서화하기 위한 **글로벌 기업 자산 관리 마스터 아키텍처**입니다.

---

## 📐 저장소 아키텍처 & 레이어 구조 (System Layer)

`mermaid
flowchart TD
    subgraph Layer1 [1. 마스터 표준 문서 Layer]
        A1[Universal_Enterprise_Asset_Management_Framework.md] --> A2[Real_Estate_Management_Rules.md]
        A2 --> A3[Antigravity_CLI_GoogleDrive_Setup_Guide.md]
    end

    subgraph Layer2 [2. 코어 파이썬 처리 엔진 Layer - src/]
        B1[config.py] --> B2[db_manager.py]
        B2 --> B3[real_estate_engine.py]
        B2 --> B4[vehicle_engine.py]
        B2 --> B5[insurance_engine.py]
        B3 & B4 & B5 --> B6[template_exporter.py]
    end

    subgraph Layer3 [3. 마스터 서식 템플릿 Layer - templates/]
        C1[Master_Corporate_Contract_Note_Overhead_Template.md]
        C2[Corporate_Real_Estate_Contract_Note_Template.md]
        C3[Corporate_Vehicle_Contract_Note_Template.md]
        C4[Corporate_Insurance_Contract_Note_Template.md]
        C5[External_Audit_IPO_Asset_Register_Template.md]
    end

    Layer1 --> Layer2
    Layer2 --> Layer3
    Layer3 --> D[기업별 실행 저장소: FoxConnect-AX 등]
`

---

## 📁 마스터 저장소 구획 및 파일 아키텍처 (Directory Tree)

`	ext
AX-Work/
├── 📄 README.md                                      <-- [총괄] 마스터 저장소 아키텍처 및 운용 가이드
├── 📄 Universal_Enterprise_Asset_Management_Framework.md <-- [표준 01] 범용 기업 자산관리 통합 프레임워크
├── 📄 Real_Estate_Management_Rules.md               <-- [표준 02] 부동산 자산 세부 실무 및 보고서 규칙
├── 📄 Antigravity_CLI_GoogleDrive_Setup_Guide.md     <-- [표준 03] CLI & 구글 드라이브 연동 가이드
├── 🚀 main.py                                         <-- [실행] 범용 마스터 파이프라인 실행 엔드포인트
├── ⚙️ mcp.json.template                              <-- [설정] Google Drive MCP 설정 샘플
├── 🔒 .gitignore                                      <-- [보안] 민감 정보 및 개별 기업 데이터 커밋 방지
│
├── 🧠 src/                                            <-- [엔진] 자산 파싱 및 DB 관리 코어 모듈
│   ├── config.py                                   <-- 동적 환경 설정 및 시스템 상수 관리
│   ├── db_manager.py                               <-- 범용 SQLite 데이터베이스 CRUD 관리 모듈
│   ├── folder_structure_engine.py                  <-- 2단계 마스터 폴더 계층 자동 생성 엔진
│   ├── real_estate_engine.py                       <-- 부동산 계약서 AI 파싱 및 이력 관리 엔진
│   ├── vehicle_engine.py                           <-- 법인 차량 계약 파싱 및 관리 엔진
│   ├── insurance_engine.py                         <-- 법인 보험 계약 파싱 및 관리 엔진
│   └── template_exporter.py                        <-- DOCX / XLSX 최종 보고서 생성 엔진
│
└── 📑 templates/                                      <-- [템플릿] 마스터 표준 서식 보관소
    ├── Master_Corporate_Contract_Note_Overhead_Template.md  <-- 마스터 통합 임대차 계약 관리 노트 서식
    ├── Corporate_Real_Estate_Contract_Note_Template.md      <-- 부동산 자산 개별 계약 노트 서식
    ├── Corporate_Vehicle_Contract_Note_Template.md          <-- 법인 차량 개별 계약 노트 서식
    ├── Corporate_Insurance_Contract_Note_Template.md        <-- 법인 보험 개별 계약 노트 서식
    ├── External_Audit_IPO_Asset_Register_Template.md        <-- 외부감사 및 IPO 제출용 자산 대장 서식
    ├── generate_docx_note.py                                 <-- 워드 노트 자동 생성 스크립트
    └── README.md                                             <-- 템플릿 사용법 및 수록 안내
`

---

## 📑 핵심 모듈 & 표준 문서 역할 분담표

| 분류 | 파일 / 모듈명 | 주요 역할 & 기능 설명 | 비고 |
| :--- | :--- | :--- | :--- |
| **표준 문서** | Universal_Enterprise_..._Framework.md | 자산 분류, 2단계 폴더 계층, 파일명 명명 규칙([파일순번]_[건물명]...) 정립 | 마스터 준수 |
| **표준 문서** | Real_Estate_Management_Rules.md | 부동산 계약 실무 규칙, A안(실무용)/B안(감사용)/C안(커스텀) 리포트 정의 | 마스터 준수 |
| **파이썬 엔진** | src/db_manager.py | 부동산, 차량, 보험 마스터 테이블 통합 관리 및 데이터 무결성 보장 | Core Engine |
| **파이썬 엔진** | src/real_estate_engine.py | 부동산 임대차 계약서 PDF AI OCR 파싱 및 릴레이션 매핑 | Core Engine |
| **파이썬 엔진** | src/template_exporter.py | 파싱 완료된 메타데이터를 Word(.docx) / Excel(.xlsx)로 내보내기 | Core Engine |
| **마스터 템플릿** | 	emplates/*.md | 외부감사/IPO 표준을 충족하는 종합 계약 노트 및 자산 대장 양식 | Standard Template |

---

## 🔄 마스터 표준 피드백 이관 (Feedback Loop) 프로세스

`	ext
[ FoxConnect-AX 실무 수행 ]
          │
          ▼ (실무 중 우수 양식 / 파이썬 코드 / 프롬프트 발굴)
[ AI 에이전트 비교 분석 ] ──> 기존 AX-Work 파일과의 차이점 및 범용성 분석 보고서 작성
          │
          ▼ (관리자 승인)
[ AX-Work 마스터 반영 ] ──> [전면 교체] 또는 [하단 섹션 추가]를 통해 마스터 표준 고도화
`

---

## 🛡️ 보안 및 오픈소스 마스터 준수
* 실제 기업의 계약서, 개인정보, .csv, .xlsx, .docx, .pdf, .db 실행 데이터는 .gitignore에 의해 본 저장소에서 전면 자동 제외됩니다.
* 본 저장소는 표준 청사진 역할을 수행하므로, 모든 커밋 및 변경 사항은 관리자 검토 후 반영됩니다.
