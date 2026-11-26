# -*- coding: utf-8 -*-
"""一次修好: (1) install.ps1 加 fallback + BOM  (2) publish_github.py 改成 repo 根結構"""
from pathlib import Path

root = Path(r"E:\1shopee\_github_pack")

# ── (1) install.ps1 ──
p = root / "install.ps1"
t = p.read_text(encoding="utf-8-sig")
old = '$skillDir = Join-Path $src "m1688-selection-ops"'
new = ('$skillDir = Join-Path $src "m1688-selection-ops"\n'
       '# fallback: 若使用者只下載了技能資料夾 (本檔與技能同層)\n'
       'if (-not (Test-Path $skillDir)) { $skillDir = $src }')
if old in t and "fallback: 若使用者只下載了技能資料夾" not in t:
    t = t.replace(old, new)
    p.write_text(t, encoding="utf-8-sig")
    print("  install.ps1: fallback 已加 + UTF-8 BOM")
else:
    print("  install.ps1: 已處理或未匹配 (保持原樣)")

# ── (2) publish_github.py: 改成上傳 repo 根結構 ──
p2 = root / "publish_github.py"
t2 = p2.read_text(encoding="utf-8")
t2 = t2.replace(
    'PKG = Path(r"E:\\1shopee\\_github_pack\\m1688-selection-ops")',
    'PKG = Path(r"E:\\1shopee\\_github_pack")   # repo 根 (含 README/install.ps1/m1688-selection-ops/)')
old2 = """    files = []
    for p in sorted(PKG.rglob("*")):
        if p.is_dir():
            continue
        rel = p.relative_to(PKG).as_posix()
        files.append((rel, p))"""
new2 = """    files = []
    skip_names = {"publish_github.py"}
    for p in sorted(PKG.rglob("*")):
        if p.is_dir():
            continue
        if p.name in skip_names or p.suffix == ".zip":
            continue
        rel = p.relative_to(PKG).as_posix()
        files.append((rel, p))"""
assert old2 in t2, "publish pattern not found"
t2 = t2.replace(old2, new2)
p2.write_text(t2, encoding="utf-8")
print("  publish_github.py: 改為 repo 根結構 (排除 publish_github.py 與 zip)")
