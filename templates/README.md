# FoxConnect Real Estate Template Storage
이 폴더(`templates/`)는 사용자가 정의한 보고서, 엑셀 대장, 양식 템플릿(.xlsx, .docx, .html 등)을 저장하는 공간입니다.

## 📁 템플릿 구조
- `real_estate_yearly_template.xlsx`: 연도별 현황 보고서 템플릿
- `snapshot_report_template.xlsx`: 시점별(12월 31일 기준) 스냅샷 보고서 템플릿
- `contract_summary_template.docx`: 계약 요약 서식 템플릿

## 💡 사용법
1. 원하는 서식 템플릿 파일을 이 폴더에 넣습니다.
2. 엔진(`src/template_exporter.py`)이 템플릿을 자동으로 읽어서 데이터와 결합하여 `output/` 폴더에 최종 보고서를 생성합니다.
