import urllib.request
import json
import base64

GITHUB_USERNAME = "barista9980-cloud"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
REPO_NAME = "AX-Work"

headers = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json",
    "User-Agent": "Antigravity-Agent"
}

with open("PREVIOUS_README.md", "r", encoding="utf-8") as f:
    prev_content = f.read()

merged_readme = f"""{prev_content}

---

# 🚀 Antigravity CLI & Google Drive Integration Starter Kit

Antigravity CLI 및 Google Drive MCP 연동을 위한 **자동화 기초 스타터 키트 저장소**입니다.

## 📌 주요 문서 및 설치 가이드
- 📄 **상세 세팅 가이드**: [Antigravity_CLI_GoogleDrive_Setup_Guide.md](Antigravity_CLI_GoogleDrive_Setup_Guide.md)
- 📄 **부동산 자산 세부 지침**: [real-estate.md](real-estate.md)

## 📁 저장소 구성
```text
.
├── README.md                                    <-- 메인 가이드라인 & 스타터 키트
├── Antigravity_CLI_GoogleDrive_Setup_Guide.md <-- 전체 5단계 환경구축 가이드
├── real-estate.md                               <-- 부동산 자산 관리 세부 지침
├── mcp.json.template                           <-- Google Drive MCP 설정 샘플
├── .gitignore                                   <-- 비밀키 및 개인 데이터 업로드 방지 규칙
├── src/                                         <-- DB 및 리포트 처리 엔진 모듈
└── templates/                                   <-- 보고서 양식 템플릿 저장소
```

## 🔒 보안 주의사항
`mcp.json` 파일에는 Google Cloud OAuth 비밀키(`Client Secret`)가 포함되므로 깃허브에 직접 커밋하지 마세요. 대신 `mcp.json.template`을 복사하여 `mcp.json`으로 이름을 바꾼 뒤 사용하세요.
"""

def get_file_sha(path):
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/contents/{path}"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            return data.get("sha")
    except Exception:
        return None

def update_readme():
    sha = get_file_sha("README.md")
    b64_content = base64.b64encode(merged_readme.encode('utf-8')).decode('utf-8')
    
    payload = json.dumps({
        "message": "Merge original AI guidelines and Antigravity setup guide into README.md",
        "content": b64_content,
        "sha": sha
    }).encode('utf-8')
    
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/contents/README.md"
    req = urllib.request.Request(url, data=payload, headers=headers, method="PUT")
    
    try:
        with urllib.request.urlopen(req) as resp:
            print("Successfully merged and restored previous README.md content!")
    except Exception as e:
        print(f"Error updating README: {e}")

if __name__ == "__main__":
    update_readme()
