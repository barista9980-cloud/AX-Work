# 🚀 Antigravity CLI & Google Drive 연동 환경 구축 종합 가이드 (Setup Guide)

본 가이드는 **Antigravity CLI** 환경에서 **Google Drive API 및 MCP(Model Context Protocol)**를 연동하여 PC 및 구글 드라이브 상에서 AI 자동화 작업을 수행하기 위한 **범용 기초 세팅 가이드 문서**입니다. 

나중에 새로운 프로젝트나 PC에서 처음부터 다시 환경을 구축할 때 이 순서대로 진행하시면 됩니다.

---

## 📌 전체 세팅 프로세스 요약 (5단계)

1. **기초 환경 설치** (Python, Node.js, Antigravity CLI)
2. **Google Cloud Console OAuth 2.0 인증 정보 발급** (Client ID & Client Secret)
3. **MCP 설정 파일 (`mcp.json`) 구성**
4. **Google Drive 데스크톱 앱 연동 및 자동 변환 옵션 설정**
5. **프로젝트 아키텍처 및 커스텀 스킬(`SKILL.md`) 세팅**

---

## 🛠️ 1단계: 필수 환경 설치 및 버전 확인

터미널(PowerShell / Command Prompt)을 열고 아래 프로그램들이 정상적으로 작동하는지 확인합니다.

```bash
# 1. 파이썬 설치 및 버전 확인
python --version

# 2. Node.js / npx 지원 확인 (MCP 서버 실행용)
npx --version

# 3. Antigravity CLI 실행 확인
agy
```

*파이썬 라이브러리가 필요한 경우 아래 명령어로 필수 패키지를 설치합니다:*
```bash
python -m pip install openpyxl pypdf
```

---

## 🔑 2단계: Google Cloud Console OAuth 2.0 인증키 발급

Google Drive API를 AI 및 MCP와 안전하게 연동하기 위한 인증 정보 발급 절차입니다.

1. **Google Cloud Console 접속 및 로그인**: [https://console.cloud.google.com/](https://console.cloud.google.com/)
2. **프로젝트 생성**:
   - 상단 프로젝트 선택 메뉴 > `[새 프로젝트]` 클릭 ➔ 프로젝트 이름 입력 후 생성
3. **Google Drive API 활성화**:
   - 왼쪽 메뉴 `[API 및 서비스]` > `[라이브러리]` 이동
   - 검색창에 **"Google Drive API"** 검색 후 `[사용 설정(Enable)]` 클릭
4. **OAuth 동의 화면 (Consent Screen) 설정**:
   - `[API 및 서비스]` > `[OAuth 동의 화면]` 이동
   - User Type: `외부(External)` 선택 후 생성
   - 앱 이름, 사용자 지원 이메일 입력 후 저장
   - `[테스트 사용자(Test Users)]` 탭에서 작업에 사용할 구글 이메일 주소 추가
5. **OAuth 2.0 클라이언트 ID 생성**:
   - `[API 및 서비스]` > `[사용자 인증 정보]` > `[+ 사용자 인증 정보 만들기]` > `[OAuth 클라이언트 ID]` 클릭
   - 애플리케이션 유형: **데스크톱 앱 (Desktop app)** 선택
   - 이름 입력 후 `[생성]`
   - 화면에 출력되는 **클라이언트 ID (Client ID)** 및 **클라이언트 보안 비밀 (Client Secret)** 복사

---

## ⚙️ 3단계: MCP 설정 파일 (`mcp.json`) 작성

발급받은 인증 정보를 바탕으로 프로젝트 루트 디렉터리에 `mcp.json` 파일을 작성합니다.

```json
{
  "mcpServers": {
    "gdrive": {
      "command": "npx",
      "args": [
        "-y",
        "@modelcontextprotocol/server-gdrive"
      ],
      "env": {
        "GDRIVE_CLIENT_ID": "발급받은_CLIENT_ID_입력",
        "GDRIVE_CLIENT_SECRET": "발급받은_CLIENT_SECRET_입력"
      }
    }
  }
}
```

---

## ☁️ 4단계: Google Drive 데스크톱 앱 연동 & 구글 시트 자동 변환 설정

### 1) Google Drive 데스크톱 앱 설치
- [구글 드라이브 데스크톱 다운로드 페이지](https://www.google.com/drive/download/)에서 설치 파일 다운로드 및 로그인.
- 설치 완료 시 PC 파일 탐색기에 `G:\내 드라이브` (또는 지정된 드라이브 문자로) 자동 마운트됨.

### 2) 구글 드라이브 웹 자동 변환 설정 (추천 1초 설정)
엑셀/CSV 파일이 구글 드라이브에 생성될 때 수동 변환 없이 **자동으로 구글 시트(.gsheet)로 변환**되도록 설정합니다.
1. [Google Drive Web](https://drive.google.com) 접속
2. 우측 상단 **⚙️ [설정] 아이콘 ➔ [설정] 클릭**
3. **[일반]** 탭 ➔ **"업로드된 파일을 Google 문서 편집기 형식으로 변환"** 체크박스 선택

---

## 🏗️ 5단계: 표준 프로젝트 아키텍처 및 커스텀 스킬 세팅

### 1) 권장 프로젝트 폴더 구조
반복적인 데이터 처리 및 보고서 생성을 위해 작업 디렉터리를 모듈화합니다.

```text
작업_프로젝트_폴더/
 ├── 📁 data/                           <-- SQLite DB 및 기초 데이터 저장소
 ├── 📁 templates/                      <-- 보고서 양식/서식 템플릿 (.xlsx, .docx)
 ├── 📁 output/                         <-- 파이썬/AI가 생성한 출력 보고서
 ├── 📁 src/                            <-- 코어 처리 파이썬 스크립트 (DB 모듈, 파서 모듈)
 ├── 📁 .gemini/skills/                <-- Antigravity Custom Skill 저장소
 └── 📄 mcp.json                        <-- 구글 드라이브 MCP 설정
```

### 2) Antigravity Custom Skill (`SKILL.md`) 세팅
반복되는 업무 워크플로우를 AI가 기억하게 하려면 프로젝트 내 `.gemini/skills/<스킬이름>/SKILL.md` 경로에 지침서를 만듭니다.

#### `SKILL.md` 작성 예시:
```yaml
---
name: my-automation-skill
description: 특정 업무 자동화 처리 및 구글 드라이브 보고서 자동 생성 스킬
---

# 스킬 가이드 내용 작성
- 실행 파이썬 스크립트 지정
- 데이터 처리 및 출력 규칙 명시
```

### 3) 대화 세션 영구 보존 (`/learn`)
터미널 CLI 대화 도중 중요한 규칙이나 성공한 패턴을 AI에 영구 반영하고 싶을 때는 대화창에 아래 슬래시 커맨드를 입력합니다:
```bash
/learn
```

---

## 💡 자주 쓰는 CLI 명령어 요약

- `agy`: Antigravity CLI 실행
- `/plan`: 대규모 작업 전 계획 수립 모드
- `/schedule`: 정기 실행 또는 알림 타이머 설정
- `/learn`: 성공한 워크플로우 영구 저장
- `/exit`: CLI 종료
