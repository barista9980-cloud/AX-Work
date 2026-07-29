---
name: real-estate-asset-manager
description: FoxConnect 부동산 계약서 파싱, 자산 데이터베이스 관리, 법인 사명 통합 및 특정 시점(12월 31일 등) 기준 자산 현황 보고서(엑셀/CSV/구글시트) 자동 생성 스킬
---

# 🏢 FoxConnect 부동산 자산 관리 및 보고서 자동 생성 스킬 (Real Estate Asset Manager)

이 스킬은 폭스커넥트(구 폭스에듀)의 부동산 계약서 PDF 파싱, 자산 데이터베이스 쿼리, 시점별 현황 보고서 자동 생성 및 구글 드라이브 연동을 수행합니다.

---

## 🎯 주요 기능 및 명령어 트리가거

사용자가 아래와 같은 요청을 할 때 이 스킬을 활성화하여 즉시 실행합니다:
1. **"YYYY-MM-DD 기준 부동산 현황 보고서 만들어줘"** (예: 2025년 12월 31일 기준, 2024년 12월 31일 기준 등)
2. **"계약 시작일 기준으로 보고서 작성해줘"**
3. **"새로운 부동산 계약서 파싱해서 자산대장 업데이트해줘"**

---

## 🛠️ 시스템 구성 및 실행 스크립트

### 1. 코어 파일 위치
- **자산 데이터베이스**: `C:\Users\User\OneDrive\바탕 화면\업무_AX\data\real_estate_assets.db`
- **구글 드라이브 동기화 폴더**: `G:\내 드라이브\[부동산자산] FoxConnect 계약 관리\`
- **보고서 저장 위치**: `G:\내 드라이브\[부동산자산] FoxConnect 계약 관리\04_생성_보고서\`

### 2. 실행 명령어 (파이썬 모듈)

#### [기능 A] 특정 시점(YYYY-MM-DD) 기준 스냅샷 보고서 생성
```bash
python -c "from src.template_exporter import TemplateExporter; exporter = TemplateExporter(); exporter.generate_snapshot_report('YYYY-MM-DD')"
```

#### [기능 B] 구글 시트 호환 엑셀/CSV 자동 생성 및 드라이브 싱크
```bash
python export_2025_gsheets.py
```

#### [기능 C] 신규 계약서 파싱 및 DB 동기화
```bash
python parse_real_estate_docs.py
python src/db_manager.py
```

---

## 📋 규칙 및 비즈니스 로직 (Rules)

1. **법인명 통합 매핑**:
   - `㈜폭스에듀`와 `㈜폭스커넥트`는 동일 법인으로 매핑하며, `(주)폭스커넥트 [구 (주)폭스에듀]`로 표준화합니다.
2. **당사 구분 자동 분류**:
   - `당사 임차 (임차인/전차인)`: 폭스커넥트가 임차한 물건
   - `당사 전대/임대 (임대인/전대인)`: 폭스커넥트가 외부 업체에 전대한 물건
   - `당사 매매`: 자사 소유 자산
3. **출력 규칙**:
   - 모든 보고서는 UTF-8-BOM 인코딩의 CSV 및 openpyxl 스타일링이 적용된 `.xlsx`로 동시 생성되며, 구글 드라이브 `04_생성_보고서` 폴더에 즉시 연동됩니다.
