# -*- coding: utf-8 -*-
"""開啟 1688 手機版首頁, 停在首頁不操作, 並提醒使用者先手動搜尋幾個詞
用法: python prompt_manual_search.py [端口, 預設 9223] [詞1 詞2 詞3 ...]
     不帶詞就用 DEFAULT_WORDS; 帶了就用你給的 (建議給當前採集任務的關鍵詞)
"""
import sys, io, json, urllib.request
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PORT = sys.argv[1] if len(sys.argv) > 1 else "9223"
CDP = f"http://127.0.0.1:{PORT}"

# 建議使用者手動搜尋的詞 (依當前採集任務, 可用引數覆蓋)
DEFAULT_WORDS = ["儿童滑板车", "露营椅", "带轮躺椅", "行军床", "月亮椅"]
_words = sys.argv[2:]
DEFAULT_WORDS = _words if _words else DEFAULT_WORDS

with sync_playwright() as p:
    b = p.chromium.connect_over_cdp(CDP)
    ctx = b.contexts[0]

    # 只關明確的桌機版分頁
    for pg in list(ctx.pages):
        u = pg.url or ""
        if ("work.1688.com" in u) or ("www.1688.com" in u):
            try:
                pg.close()
            except Exception:
                pass

    page = None
    for pg in ctx.pages:
        if "m.1688.com" in (pg.url or ""):
            page = pg
            break
    if page is None:
        page = ctx.new_page()

    page.set_viewport_size({"width": 414, "height": 896})
    page.goto("https://m.1688.com/", wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(2500)
    page.bring_to_front()

    nick = lid = ""
    for c in ctx.cookies():
        if c.get("name") == "nick": nick = c.get("value", "")
        elif c.get("name") == "lid": lid = c.get("value", "")

    print("=" * 58)
    print(f"  1688 手機版已開啟 (端口 {PORT})，目前停在首頁，尚未做任何搜尋")
    print(f"  登入: nick={nick or '(空)'}  lid={lid}")
    print("=" * 58)
    print("  請你在這個視窗手動搜尋以下 3-5 個詞（每個詞搜完停留 3-5 秒）:")
    for w in DEFAULT_WORDS:
        print(f"     - {w}")
    print()
    print("  搜完後回覆我一聲「好了」，我就接著開跑採集。")
    print("  理由: 真人手動搜尋是平台判定正常使用者最直接的訊號，")
    print("        做完這幾個詞，搜索配額會恢復，之後的自動化才不會立刻被擋。")
    print("=" * 58)
    b.close()
