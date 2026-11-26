# -*- coding: utf-8 -*-
"""diag_token.py — 診斷 fine-grained PAT 到底能做什麼"""
import base64
import json
import os
import sys
import urllib.error
import urllib.request

API = "https://api.github.com"
T = os.environ.get("GITHUB_TOKEN", "").strip()


def api(method, path, data=None):
    url = API + path
    body = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header("Authorization", "Bearer " + T)
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    req.add_header("User-Agent", "diag")
    if body:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=45) as r:
            txt = r.read().decode("utf-8", "replace")
            return True, r.status, (json.loads(txt) if txt else {})
    except urllib.error.HTTPError as e:
        return False, e.code, e.read().decode("utf-8", "replace")[:300]
    except Exception as e:
        return False, -1, str(e)[:200]


OWNER = "Xun209"
REPO = "1688-selection-skill"
out = []

ok, st, data = api("GET", "/user")
out.append(f"GET  /user                 -> {st}  {data.get('login') if ok else data}")

ok, st, data = api("GET", f"/repos/{OWNER}/{REPO}")
out.append(f"GET  repo                  -> {st}  {data.get('html_url') if ok else data}")
if ok:
    out.append(f"     permissions = {data.get('permissions')}")
    out.append(f"     private={data.get('private')} default_branch={data.get('default_branch')}")

ok, st, data = api("GET", f"/repos/{OWNER}/{REPO}/contents")
out.append(f"GET  contents (讀)         -> {st}  {'可讀' if ok else data}")

# 測試寫入 (Contents API)
ok, st, data = api("PUT", f"/repos/{OWNER}/{REPO}/contents/_probe.txt",
                   {"message": "probe", "content": base64.b64encode(b"probe\n").decode()})
out.append(f"PUT  contents (寫)         -> {st}  {'可寫 ✓' if ok else data}")

if ok:
    sha = data.get("content", {}).get("sha")
    ok2, st2, d2 = api("DELETE", f"/repos/{OWNER}/{REPO}/contents/_probe.txt",
                       {"message": "probe cleanup", "sha": sha})
    out.append(f"DEL  probe cleanup        -> {st2}")

ok, st, data = api("GET", f"/repos/{OWNER}/{REPO}/git/refs/heads/main")
out.append(f"GET  git/refs/heads/main   -> {st}  {'有 main 分支' if ok else data}")

ok, st, data = api("GET", "/user/repos?per_page=5")
if ok:
    out.append("GET  /user/repos (可見的 repo):")
    for r in data[:5]:
        out.append(f"     {r['full_name']}  private={r['private']}")

print("\n".join(out))
