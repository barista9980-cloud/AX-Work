import urllib.request
import json

GITHUB_USERNAME = "barista9980-cloud"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
REPO_NAME = "AX-Work"

headers = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json",
    "User-Agent": "Antigravity-Agent"
}

def check_repo_commits():
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/commits"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            commits = json.loads(resp.read().decode('utf-8'))
            print(f"Total commits in {REPO_NAME}: {len(commits)}")
            for c in commits:
                commit_info = c["commit"]
                print(f"Commit SHA: {c['sha'][:7]} | Message: {commit_info['message']} | Date: {commit_info['committer']['date']}")
    except Exception as e:
        print(f"Error fetching commits: {e}")

def check_all_repo_tree():
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/git/trees/main?recursive=1"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            tree = data.get("tree", [])
            print(f"\nAll files in repo currently ({len(tree)} items):")
            for item in tree:
                print(f"  - {item['path']} ({item['type']})")
    except Exception as e:
        print(f"Error fetching tree: {e}")

if __name__ == "__main__":
    check_repo_commits()
    check_all_repo_tree()
