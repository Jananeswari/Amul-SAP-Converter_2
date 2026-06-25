"""
GitHub Persistence Helper
Pushes the master file (and its snapshot) back to the GitHub repo that
Streamlit Cloud deploys from, so edits made through the website survive
the app sleeping or redeploying — instead of being lost the moment
Streamlit Cloud's disk resets.

Requires three Streamlit secrets to be configured (see SETUP_GUIDE.md):
  GITHUB_TOKEN   - a personal access token with 'repo' (or fine-grained
                   'contents: read/write') permission on the target repo
  GITHUB_REPO    - "owner/repo-name", e.g. "Jananeswari/Amul-SAP-Converter"
  GITHUB_BRANCH  - usually "main"

If these secrets are not configured, push_file_to_github simply returns
a clear "not configured" result instead of raising — so the rest of the
app (and local/offline use) keeps working without GitHub at all.
"""

import base64
import requests


def is_github_configured(secrets) -> bool:
    return all(k in secrets for k in ("GITHUB_TOKEN", "GITHUB_REPO", "GITHUB_BRANCH"))


def push_file_to_github(local_path: str, repo_path: str, commit_message: str, secrets) -> dict:
    """
    Pushes the file at local_path to repo_path inside the configured
    GitHub repo, creating or updating it as needed (GitHub's Contents
    API requires the current file's SHA to update an existing file).

    Returns {"success": bool, "message": str}
    """
    if not is_github_configured(secrets):
        return {"success": False, "message": "GitHub is not configured (missing secrets) — change saved locally only."}

    token = secrets["GITHUB_TOKEN"]
    repo = secrets["GITHUB_REPO"]
    branch = secrets["GITHUB_BRANCH"]

    api_url = f"https://api.github.com/repos/{repo}/contents/{repo_path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
    }

    try:
        with open(local_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("utf-8")

        # Look up current file SHA, if it already exists on this branch
        get_resp = requests.get(api_url, headers=headers, params={"ref": branch}, timeout=15)
        sha = get_resp.json().get("sha") if get_resp.status_code == 200 else None

        payload = {
            "message": commit_message,
            "content": content_b64,
            "branch": branch,
        }
        if sha:
            payload["sha"] = sha

        put_resp = requests.put(api_url, headers=headers, json=payload, timeout=15)

        if put_resp.status_code in (200, 201):
            return {"success": True, "message": f"Pushed to GitHub ({repo}, {repo_path})."}
        else:
            detail = put_resp.json().get("message", put_resp.text)
            return {"success": False, "message": f"GitHub push failed ({put_resp.status_code}): {detail}"}

    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"GitHub push failed (network error): {e}"}
    except Exception as e:
        return {"success": False, "message": f"GitHub push failed: {e}"}
