import os
import json
import urllib.request
import urllib.error
import base64

GITHUB_USERNAME = "barista9980-cloud"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
REPO_NAME = "AX-Work"
BASE_DIR = os.path.dirname(__file__)

headers = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json",
    "User-Agent": "Antigravity-Agent"
}

def ensure_repository_exists():
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            print(f"Repository '{REPO_NAME}' exists.")
            return True
    except urllib.error.HTTPError as e:
        if e.code == 404:
            print(f"Repository '{REPO_NAME}' not found. Creating repository...")
            create_url = "https://api.github.com/user/repos"
            payload = json.dumps({
                "name": REPO_NAME,
                "description": "Antigravity CLI & Google Drive Work Starter Kit",
                "private": False
            }).encode('utf-8')
            
            create_req = urllib.request.Request(create_url, data=payload, headers=headers, method="POST")
            try:
                with urllib.request.urlopen(create_req) as create_resp:
                    print(f"Successfully created repository '{REPO_NAME}'.")
                    return True
            except Exception as create_err:
                print(f"Failed to create repo: {create_err}")
                return False
        else:
            print(f"HTTP Error checking repo: {e}")
            return False

def get_file_sha(path_in_repo):
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/contents/{path_in_repo}"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            return data.get("sha")
    except urllib.error.HTTPError:
        return None

def upload_file_to_repo(local_path, path_in_repo, commit_message):
    if not os.path.exists(local_path):
        print(f"Local file missing: {local_path}")
        return False

    with open(local_path, "rb") as f:
        content_bytes = f.read()

    b64_content = base64.b64encode(content_bytes).decode('utf-8')
    sha = get_file_sha(path_in_repo)

    payload_dict = {
        "message": commit_message,
        "content": b64_content
    }
    if sha:
        payload_dict["sha"] = sha

    payload = json.dumps(payload_dict).encode('utf-8')
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/contents/{path_in_repo}"
    
    req = urllib.request.Request(url, data=payload, headers=headers, method="PUT")
    try:
        with urllib.request.urlopen(req) as resp:
            print(f"Successfully uploaded {path_in_repo} to GitHub!")
            return True
    except Exception as e:
        print(f"Error uploading {path_in_repo}: {e}")
        return False

def main():
    if not ensure_repository_exists():
        return

    files_to_upload = [
        ("README.md", "README.md", "Add README.md starter guide"),
        ("Antigravity_CLI_GoogleDrive_Setup_Guide.md", "Antigravity_CLI_GoogleDrive_Setup_Guide.md", "Add Antigravity CLI Setup Guide"),
        ("mcp.json.template", "mcp.json.template", "Add MCP configuration template"),
        (".gitignore", ".gitignore", "Add gitignore rules for security"),
        ("src/db_manager.py", "src/db_manager.py", "Add DB manager engine module"),
        ("src/template_exporter.py", "src/template_exporter.py", "Add Template Exporter module"),
        ("templates/README.md", "templates/README.md", "Add templates directory README")
    ]

    for local_rel, repo_rel, msg in files_to_upload:
        full_local = os.path.join(BASE_DIR, local_rel)
        upload_file_to_repo(full_local, repo_rel, msg)

    print("\nAll files successfully committed to GitHub repository!")

if __name__ == "__main__":
    main()
