# -*- coding: utf-8 -*-
"""probe_preflight.py — 開跑前先探一次 mtop, 確認帳號沒被風控擋住

★ 2026-11-06 重大修正: 原本用 urllib 發「裸請求」→ 沒有瀏覽器 TLS 指紋/
  client hints / 完整 cookie → 1688 對非瀏覽器客戶端本來就回 FAIL_SYS_USER_VALIDATE
  → 這是「假陽性風控」, 會讓人白白停工 (2026-11-06 因此誤判三次)
  → 現在改成「在瀏覽器頁面 context 內 fetch」, 與採集器完全同一條路徑

exit 0 = 正常 (ret 裡有 SUCCESS)
exit 2 = 真風控 (在瀏覽器 context 內被拒絕 / 被導到 punish 驗證頁)
exit 3 = 網路、分頁或登入異常 (≠ 風控)
"""
import sys
import time

sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

KW = sys.argv[1] if len(sys.argv) > 1 else "宠物碗斜口防噎"
_kw_arg = sys.argv[1] if len(sys.argv) > 1 else None
sys.argv = [sys.argv[0]]

import collect_m1688_v3_api as C
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    b = p.chromium.connect_over_cdp(C.CDP_URL)
    ctx = b.contexts[0]

    # 1) 找一個 m.1688.com 的分頁 (手機端 origin); 沒有就自己開一個
    page = None
    for pg in ctx.pages:
        if "m.1688.com" in (pg.url or ""):
            page = pg
            break
    if page is None:
        page = ctx.new_page()
        try:
            page.set_viewport_size({"width": 414, "height": 896})
        except Exception:
            pass
        try:
            page.goto(C.COLLECT_URL, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(3000)
        except Exception as e:
            print(f"[PREFLIGHT] 無法開啟手機端分頁: {str(e)[:90]}")
            sys.exit(3)

    url_now = page.url or ""
    if "m.1688.com" not in url_now:
        print(f"[PREFLIGHT] 分頁不在手機端: {url_now[:90]}")
        print("           → Chrome 必須用手機 UA 啟動, 否則 m.1688.com 會被導到桌機版")
        sys.exit(3)

    # 2) 取 token
    token = ""
    for c in ctx.cookies():
        if c["name"] == "_m_h5_tk":
            token = c["value"].split("_")[0]
    if not token:
        print("[PREFLIGHT] 拿不到 _m_h5_tk, 需要重新登入")
        b.close()
        sys.exit(3)

    # 3) 在頁面 context 內發 (與採集器同一路徑)
    u = C.build_mtop_url(token, KW, 1)
    try:
        r = page.evaluate("""async (u) => {
            const ctrl = new AbortController();
            const timer = setTimeout(() => ctrl.abort(), 35000);
            try {
                const resp = await fetch(u, {credentials: 'include', signal: ctrl.signal});
                return {ok: resp.ok, status: resp.status, text: await resp.text()};
            } catch (e) {
                return {ok: false, status: 0, text: 'FETCH_ERR ' + String(e).slice(0, 80)};
            } finally { clearTimeout(timer); }
        }""", u)
    except Exception as e:
        print(f"[PREFLIGHT] 頁面內請求失敗: {type(e).__name__}: {str(e)[:110]}")
        b.close()
        sys.exit(3)

    body = r.get("text") or ""
    status = r.get("status")

    # 4) 解析 (JSONP: mtopjsonp1({...}) → 要剝外殼)
    ret = ""
    punish = ""
    txt = body
    i, j = txt.find("{"), txt.rfind("}")
    if i >= 0 and j > i:
        raw = txt[i:j + 1]
        try:
            import json
            obj = json.loads(raw)
        except Exception:
            obj = None
        if obj is None:
            # JSONP 外殼: mtopjsonp1({...})
            try:
                obj = C.parse_jsonp(txt)
            except Exception:
                obj = None
        if isinstance(obj, dict):
            ret = ",".join(obj.get("ret") or [])
            punish = (obj.get("data") or {}).get("url") or ""
            found = (obj.get("data") or {}).get("found")
            if found is not None and "SUCCESS" not in ret:
                ret = ret or "SUCCESS(有結果)"
        else:
            ret = txt[:140]

    # 5) 頁面層二次確認: 有沒有被導到 punish 驗證頁 (真正的帳號級封鎖徵兆)
    page_flag = ""
    try:
        cur = page.url or ""
        if "_____tmd_____" in cur or "punish" in cur or "sec_verify" in cur:
            page_flag = cur
    except Exception:
        pass

    b.close()

# ── 判定 (2026-11-09: 用頁面狀態分類器, 區分「滑塊挑戰」與「限流」) ──
import sys as _sys
_sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
import v3_page_state as PS

_kind, _ind = PS.classify_from_body(body, "")
_dec = PS.recovery_decision(_kind)

if "SUCCESS" in ret:
    print(f"[PREFLIGHT] OK  ret={ret[:70]}  http={status}")
    sys.exit(0)

if _kind == PS.RATE_LIMITED:
    # ★ 限流不是滑塊挑戰: 退避 60-90 秒可重試, 等 40-50 分鐘更好
    print("[PREFLIGHT] 限流 (RGV587/被挤爆啦) — 不是滑塊挑戰")
    print(f"           建議: 等 {_dec['cooldown_sec']}~90 秒後重試; 更好的做法是等 40-50 分鐘讓 TTL 走完")
    print(f"           ret={ret[:120]}")
    if punish:
        print(f"           punish={punish[:170]}")
    sys.exit(9)

if _kind == PS.RISK_CHALLENGE:
    print("[PREFLIGHT] 真風控 (滑塊挑戰)")
    print(f"           ret={ret[:120]}")
    if punish:
        print(f"           punish={punish[:170]}")
    if page_flag:
        print(f"           分頁已被導到驗證頁: {page_flag[:120]}")
        print("           → 請在手機端視窗把滑塊拖過去")
    sys.exit(2)

if "FAIL_SYS_USER_VALIDATE" in ret:
    print("[PREFLIGHT] 被要求驗證 (FAIL_SYS_USER_VALIDATE), 尚非明確封鎖")
    print(f"           ret={ret[:140]}")
    print("           → 可能是 token 過期; 建議重新整理手機端分頁後再探一次")
    sys.exit(3)

print(f"[PREFLIGHT] 非預期回應 http={status} ret={ret[:160]}")
sys.exit(3)
