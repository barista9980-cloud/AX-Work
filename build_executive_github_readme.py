import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

REPO_DIR = r"C:\Users\User\OneDrive\바탕 화면\업무_AX"
readme_p = os.path.join(REPO_DIR, "README.md")

readme_content = r"""# 🏢 Enterprise AX Asset Management & Vision AI System
> **(주식회사 폭스에듀 총무 자산관리 및 비전 AI 문서자동화 프레임워크)**  
> **Repository Target**: `https://github.com/barista9980-cloud/AX-Work`  
> **Compliance Standard**: 외부감사(External Audit) 및 IPO 상장 대비 표준 자산관리 대장 체계

---

## 📌 1. 시스템 개요 (Executive Summary)

본 저장소는 **(주식회사 폭스에듀)**의 총무 자산(부동산 임대차/전대차/소유권, 법인차량 장기렌트/운용리스, 기업보험 및 배서계약)을 **Vision AI 문서 인식기술 및 1:1 독립 폴더 수칙**에 기반하여 자동 분류·정돈하고, **외부감사 및 IPO 상장 기준에 부합하는 연도별 총괄자산대장(.xlsx)**으로 자동 통합·관리하는 **AX(AI Transformation) 총무자산관리 프레임워크**입니다.

```mermaid
graph TD
    A["📄 06_자동파싱_업로드큐<br/>(신규 스캔 PDF/이미지 수집)"] --> B["🤖 Vision AI / OCR<br/>(계약서 및 약정서 정밀 파싱)"]
    B --> C1["🏢 01_부동산_자산관리<br/>(27개 1:1 독립 폴더)"]
    B --> C2["🚗 02_차량_자산관리<br/>(10대 1:1 독립 폴더)"]
    B --> C3["🛡️ 03_보험_자산관리<br/>(12개 1:1 독립 폴더)"]
    C1 --> D1["📊 부동산 총괄자산대장.xlsx<br/>(임대/전대/매매 3개 탭)"]
    C2 --> D2["📊 법인차량 총괄자산대장.xlsx<br/>(운행 8대/보증금/월세 탭)"]
    C3 --> D3["📊 기업보험 총괄자산대장.xlsx<br/>(경영인/화재/자동차 탭)"]
```

---

## 🏗️ 2. 표준 구글드라이브 디렉토리 구조 (Directory Architecture)

`G:\내 드라이브\[FoxConnect]\[총무]업무` 하위에 적용된 **고유 일련번호(01~06) 및 1:1 Flat 수평 구조**입니다:

```text
G:\내 드라이브\[FoxConnect]\[총무]업무\
│
├── 📁 00_연도별_자산현황_자료/           # 연도별 수집 원본 PDF (참고용)
│
├── 📁 01_부동산_자산관리/
│   ├── 📁 00_연도별_부동산_총괄자산대장/  # 📊 [외감_IPO대비] 부동산 마스터 엑셀 (.xlsx)
│   ├── 📁 01_임대차계약/               # 🏢 25개 임대차 1:1 독립 폴더 (01~25 continuous)
│   └── 📁 02_매매_소유권문서/           # 🏛️ 2개 매매/분양권 1:1 독립 폴더 (01~02 continuous)
│
├── 📁 02_차량_자산관리/
│   ├── 📁 00_연도별_차량_총괄자산대장/    # 📊 [외감_IPO대비] 법인차량 마스터 엑셀 (.xlsx)
│   └── 📁 01_차량계약_리스_렌트/        # 🚗 10대 차량 1:1 독립 폴더 (01~10 continuous)
│
├── 📁 03_보험_자산관리/
│   ├── 📁 00_연도별_보험_총괄자산대장/    # 📊 [외감_IPO대비] 기업보험 마스터 엑셀 (.xlsx)
│   ├── 📁 01_보험증권_및_배서계약서/     # 🛡️ 12개 기업보험 1:1 독립 폴더 (증권+배서 1:1 저장)
│   ├── 📁 02_보험금청구_사고접수/       # 🚑 사고접수 및 보험금 청구 서류
│   └── 📁 03_보험료납입_증빙/           # 💳 월/연간 보험료 납입 영수증
│
└── 📁 06_자동파싱_업로드큐/             # 📥 신규 문서 자동 분류 인테이크 큐
    ├── 📁 01_부동산_업로드대기/
    ├── 📁 02_차량_업로드대기/
    ├── 📁 03_보험_업로드대기/
    ├── 📁 04_비품_소모품_업로드대기/
    └── 📁 05_처리완료_아카이브/
```

---

## 📊 3. 외감 / IPO 대비 3대 마스터 엑셀 규격 (Master Excel Standards)

외부감사인 및 IPO 대표주관사 제출용으로 표준화된 엑셀 워크북 디자인 수칙입니다:

| 항목 | 서식 표준 규격 (Formatting Rules) |
| :--- | :--- |
| **열 여백 (Margin)** | Column A 여백 전용 설정 (`width = 3`), 표 데이터는 **Column B부터 작성** |
| **시트 배경 (Gridlines)** | 배경 격자선 제거 (`showGridLines = False`)로 깨끗한 경영진 보고 UI 구현 |
| **표 테두리 (Borders)** | 선명한 **검은색 실선 (`#000000`)** 적용으로 가독성 극대화 |
| **작성 기준일 (Date)** | `"현재"` 제거 ➔ **`작성 기준일: 2025년 12월 31일`** 단일 잔액 Snapshot 표기 |
| **숫자 방지 (No E-Notation)**| 증권번호(`0946352...`), 차량번호, 차대번호 지수표기 예방 (`@` 텍스트 서식 지정) |
| **상단 요약바 (Summary)** | Row 5에 **`작성 기준일: 2025-12-31 기준`** 유효 자산 총액 요약바 상시 배치 |

---

## 📑 4. 1:1 Flat 독립 폴더 및 Option A 파일명 규칙

모든 자산 계약은 **단 하나의 1:1 독립 폴더**와 **표준화된 파일명**을 보유합니다:

### 1) 부동산 옵션 A 파일명 규칙
- **임대차**: `[순번]_[계약유형]_[물건명]_[임대인-임차인]_(YYMMDD).pdf`
  - 예시: `01_최초임대차_강남_도곡로1길23_전층_[박재윤-㈜폭스에듀]_(241107).pdf`
- **전대차**: `[순번]_[전대차계약]_[물건명_층수]_[전대인-전차인]_(YYMMDD).pdf`
  - 예시: `05_전대차계약_강남_도곡로1길23_1층_[㈜폭스에듀-㈜에스앤에이치트레이딩]_(241101).pdf`

### 2) 법인차량 옵션 A 파일명 규칙
- `[순번]_[금융유형]_[차종]_[차량번호]_[금융사-법인]_(YYMMDD).pdf`
  - 예시: `05_운용리스_벤츠S클래스_281가8991_[하나캐피탈-㈜폭스에듀]_(220314).pdf`

### 3) 기업보험 1:1 증권 및 배서 통합 규칙
- 각 1:1 보험 폴더 내부에 **최초 증권 PDF, 배서 승인서 PDF, Word 마스터 계약노트**를 1:1로 함께 보관.
  - 예시: `12_법인차량자동차보험_카니발/` ➔ `01_최초증권`, `02_배서승인서`, `보증보험_계약관리노트_카니발.docx`

---

## 🛠️ 5. 주요 핵심 자동화 스크립트 카탈로그 (Automation Scripts)

본 저장소의 루트에 위치한 주요 Python 자동화 파이프라인 스크립트 목록입니다:

| 스크립트 파일명 | 기능 및 역할 설명 |
| :--- | :--- |
| **`build_refined_multi_tab_real_estate_excel.py`** | 🏢 부동산 외감/IPO 대비 3개 탭 마스터 엑셀 대장 자동 생성 |
| **`build_executive_vehicle_and_insurance_excels.py`** | 🚗🛡️ 법인차량 및 기업보험 외감/IPO 대비 마스터 엑셀 대장 자동 생성 |
| **`build_ipo_audit_compliant_00_registers.py`** | 📊 00_ 총괄자산대장 디렉토리 표준 CSV/Excel 통합 생성 |
| **`inspect_all_1to1_vehicle_notes_and_pdfs.py`** | 🔍 10대 차량 1:1 폴더 및 Word 노트 원본 정밀 조사를 통한 데이터 검증 |
| **`check_and_create_upload_queue_dir.py`** | 📥 `06_자동파싱_업로드큐` 5대 고유 일련번호(01~05) 수평 체계 구축 |
| **`execute_final_git_push.py`** | 🔒 보안 검수 및 GitHub `main` 브랜치 자동 커밋·푸시 실행 |

---

## 🔒 6. 보안 및 시크릿(Secret) 관리 수칙

1. **API Key 보호**: 모든 Python 코드 및 문서에서 Gemini/OpenAI 등 외부 API Key 및 토큰은 자동 검수되어 제거되며, 환경변수 또는 로컬 보안 설정 파일을 통해서만 안전하게 로드됩니다.
2. **개인정보 보호**: 계약서 스캔본 및 주민등록번호 등 민감 개인정보는 별도의 마스킹 수칙을 적용하여 안전하게 관리됩니다.

---

### 💻 Maintainer & Repository Info
- **Maintainer**: (주식회사 폭스에듀 총무팀 / AX-Work)
- **Repository**: [https://github.com/barista9980-cloud/AX-Work](https://github.com/barista9980-cloud/AX-Work)
"""

with open(readme_p, "w", encoding="utf-8") as f:
    f.write(readme_content)

print("Executive GitHub README.md updated successfully with raw string!")
