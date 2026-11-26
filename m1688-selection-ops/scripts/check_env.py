# -*- coding: utf-8 -*-
"""1688 環境檢查: 開 m.1688.com 首頁 + 確認登入 + 檢查滑塊風控 + 切到最前面
用法: python check_env.py [關鍵詞]   不帶參數則只開首頁不搜尋
"""
import sys, io, time, json
from urllib.parse import quote
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

KW = sys.argv[1] if len(sys.argv) > 1 else ""

with sync_playwright() as p:
    b = p.chromium.connect_over_cdp('http://127.0.0.1:9223')
    ctx = b.contexts[0]
    # 只關桌機版 1688 分頁 (限流主因), 絕不關全部
    for pg in list(ctx.pages):
        if '1688.com' in (pg.url or '') and 'm.1688.com' not in pg.url:
            try:
                pg.close(); print("關閉桌機版:", pg.url[:70])
            except Exception:
                pass
    tmd = [x for x in ctx.pages if '_____tmd_____' in (x.url or '')]
    if tmd:
        print("!! 滑塊風控頁還在, 已切到最前面 -> 請手動拖滑塊")
        tmd[0].bring_to_front()
        sys.exit(2)

    pg = [x for x in ctx.pages if 'm.1688.com' in x.url]
    pg = pg[0] if pg else ctx.new_page()
    pg.set_viewport_size({"width": 414, "height": 896})

    if KW:
        pg.goto('https://m.1688.com/offer_search/-6D7033.html?keywords=' + quote(KW),
                wait_until='domcontentloaded', timeout=60000)
    else:
        pg.goto('https://m.1688.com/', wait_until='domcontentloaded', timeout=60000)
    time.sleep(6)

    info = pg.evaluate("""() => {
        const c = document.cookie || '';
        const g = n => { const m = c.match(new RegExp('(?:^|;\\\\s*)'+n+'=(.*?)(?:;|$)')); return m ? m[1] : ''; };
        return { nk: g('_nk_'), lid: g('lid'), logon: g('__cn_logon__'), un: g('unb'), url: location.href };
    }""")
    print(json.dumps(info, ensure_ascii=False))

    ok = info.get("nk") == "tb966392038" and info.get("logon") == "true"
    print("登入:", "OK tb966392038" if ok else "!! 未登入, 需重新登入")

    if KW:
        cnt = 0
        for i in range(5):
            try:
                pg.mouse.wheel(0, 2500)
            except Exception:
                pass
            time.sleep(2)
            try:
                cnt = pg.evaluate("""() => { const s=new Set();
                    document.querySelectorAll('a[href*="/offer/"]').forEach(a=>{
                      const m=(a.getAttribute('href')||'').match(/offer\\/(\\d+)\\.html/); if(m) s.add(m[1]);});
                    return s.size; }""")
                print(f"  滚 {i+1} -> {cnt} 张")
            except Exception:
                print(f"  滚 {i+1} -> eval 失败")
        if '_____tmd_____' in (pg.url or ''):
            print("!! 搜尋後觸發風控, 請手動拖滑塊")
        elif cnt >= 40:
            print(f"-> 限流正常 ({cnt} 張卡), 可以開跑")
        else:
            print(f"-> 卡數偏低 ({cnt}), 可能已降級, 建議等一會再跑")

    pg.bring_to_front()
    print("已切到最前面")
