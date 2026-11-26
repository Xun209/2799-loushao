# -*- coding: utf-8 -*-
"""open_1688_only.py — 只開 1688 首頁並回報登入狀態 (不做任何搜尋/採集)
用法: python open_1688_only.py [輸出png]
"""
import sys, io, time
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)

OUT = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\夏晨茗\.mavis\sessions\mvs_210db7fd001a44fda722bd8129a2d675\workspace\open1688.png"
URL = "https://m.1688.com/?_mc=1"

with sync_playwright() as p:
    br = p.chromium.connect_over_cdp("http://127.0.0.1:9223")
    ctx = br.contexts[0]
    print(f"分頁數: {len(ctx.pages)}")
    for i, pg in enumerate(ctx.pages):
        print(f"  [{i}] {pg.url[:100]}")

    # 找既有的 1688 分頁, 沒有就用第一個; 都沒有就開新的
    target = None
    for pg in ctx.pages:
        if "m.1688.com" in pg.url:
            target = pg
            break
    if target is None:
        if ctx.pages:
            target = ctx.pages[0]
        else:
            target = ctx.new_page()
        print("-> 導向 1688 首頁")
        try:
            target.goto(URL, wait_until="domcontentloaded", timeout=30000)
        except Exception as e:
            print(f"   goto 例外: {str(e)[:120]}")
        time.sleep(4)
    else:
        print("-> 沿用既有 1688 分頁 (不重新導頁)")

    print(f"目前 URL: {target.url[:120]}")
    try:
        print(f"title: {target.title(timeout=10000)}")
    except Exception as e:
        print(f"title 取不到: {str(e)[:80]}")

    # 登入 cookie
    try:
        cookies = ctx.cookies("https://m.1688.com")
        names = {c["name"]: c["value"] for c in cookies}
        print("--- cookie ---")
        for k in ("lid", "_m_h5_tk", "cookie2", "unb", "cna", "isg", "lgc"):
            if k in names:
                v = names[k]
                print(f"  {k} = {v[:60]}{'...' if len(v) > 60 else ''}")
        print(f"  lid 存在 = {'lid' in names}")
    except Exception as e:
        print(f"cookie 取不到: {str(e)[:120]}")

    try:
        target.screenshot(path=OUT, timeout=20000)
        print(f"截圖 OK -> {OUT}")
    except Exception as e:
        print(f"截圖失敗: {str(e)[:120]}")
