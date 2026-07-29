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

def get_file_at_commit(sha, path):
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/contents/{path}?ref={sha}"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            content_b64 = data.get("content", "")
            return base64.b64decode(content_b64).decode('utf-8')
    except Exception as e:
        return None

if __name__ == "__main__":
    prev_sha = "0a4036b"
    prev_readme = get_file_at_commit(prev_sha, "README.md")
    if prev_readme:
        with open("PREVIOUS_README.md", "w", encoding="utf-8") as f:
            f.write(prev_readme)
        print("Successfully saved PREVIOUS_README.md in UTF-8 format.")
