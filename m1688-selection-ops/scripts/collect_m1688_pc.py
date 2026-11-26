# -*- coding: utf-8 -*-
"""collect_m1688_pc.py — 寵物餵食 (或任何品類) 的「PC 模式」採集器

★ 為什麼要有這條路 (2026-11-09):
  手機端 mtop 搜尋接口 (`mtop.1688.wap.seo.offer.get`) 的風險配額會被限流,
  而 PC 端搜尋頁用的是**另一個接口**:
      mtop.relationrecommend.wirelessrecommend.recommend (appId 32517, v=2.0, jsv=2.7.4)
  → 不同接口 = 不同配額桶; 加上 PC 端可用**另一個帳號** → 兩條路互不干擾。

★ 做法 (與 GitHub superjack2050/1688-cli 的 search-capture.ts 同構):
  **不自己組 payload** → 讓真實頁面發請求 → 用 page.on('response') 攔回應。
  → 請求 100% 是瀏覽器自己發的, 沒有任何「機器味」。

★ 關鍵細節:
  1. 關鍵字必須用 **GBK** 編碼 (PC 頁 charset=GBK)。用 UTF-8 送會變亂碼且只撈到不相干的結果
     (實測: GBK → found=348 / 60 筆每頁; UTF-8 → found=8 / 4 筆且標題是乱码)
  2. 回應結構: data.data.OFFER.items[].data
  3. 欄位比手機 API 更齊: 已售N+件 / 回頭率% / 實力工廠 / 深度驗廠 / 產地 / 服務標籤

用法:
  python collect_m1688_pc.py <start> <end> <tag> <詞表> [CDP埠]
  例: python collect_m1688_pc.py 34 53 pc1 petfeed 9226
"""
from __future__ import annotations

import io
import json
import os
import random
import re
import sys
import time
import urllib.parse
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

# 先擺好 argv 再 import → 被 import 的採集器會取到正確的品類/標籤
START = int(sys.argv[1]) if len(sys.argv) > 1 else 1
END = int(sys.argv[2]) if len(sys.argv) > 2 else 20
TAG = sys.argv[3] if len(sys.argv) > 3 else "pc1"
KWSET = sys.argv[4] if len(sys.argv) > 4 else "petfeed"
CDP_PORT = sys.argv[5] if len(sys.argv) > 5 else "9226"
# PC 通道用獨立帳號 → 鎖與每日預算獨立 (2026-11-10)
os.environ["M1688_LOCK_SUFFIX"] = "_pc"
os.environ["M1688_STATE_SUFFIX"] = "_pc"
sys.argv = [sys.argv[0], str(START), str(END), TAG, KWSET]

import collect_m1688_v3_api as C  # noqa: E402
import v3_page_state as PS  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

C.log = lambda *a, **k: print(*a, **k)
LOG = C.log

PC_SEARCH = "https://s.1688.com/selloffer/offer_search.htm?keywords={enc}"
MAX_PAGES = int(__import__("os").environ.get("M1688_MAX_PAGES", "3"))
GAP_MIN = float(__import__("os").environ.get("M1688_KW_GAP_MIN", "45"))
GAP_MAX = float(__import__("os").environ.get("M1688_KW_GAP_MAX", "75"))
PAGE_SETTLE_MS = int(__import__("os").environ.get("M1688_PC_SETTLE_MS", "4500"))
done_file = None
kws = []


def gbk_encode(kw: str) -> str:
    """PC 頁要 GBK; 編不出來的字丟掉 (只保留可編碼部分)"""
    try:
        return urllib.parse.quote(kw.encode("gbk"))
    except UnicodeEncodeError:
        return urllib.parse.quote(kw.encode("gbk", errors="ignore"))


def clean_title(t: str) -> str:
    t = str(t or "")
    t = re.sub(r"</?font[^>]*>", "", t)
    t = re.sub(r"<[^>]+>", "", t)
    return t.replace("&nbsp;", " ").strip()


def parse_cn_count(s):
    """「已售7600+件」「全网1.9万+件」→ 7600 / 19000"""
    if not s:
        return None
    t = str(s).replace(",", "").replace("+", "")
    m = re.search(r"([\d.]+)\s*([万亿wWkK])?", t)
    if not m:
        return None
    try:
        v = float(m.group(1))
    except ValueError:
        return None
    unit = (m.group(2) or "").lower()
    if unit in ("万", "w"):
        v *= 10000
    elif unit in ("亿",):
        v *= 100000000
    elif unit in ("k",):
        v *= 1000
    return int(v)


def pc_offer_to_item(d):
    """PC 回應的 data 節點 → 內部 item (欄位口徑對齊手機版)"""
    oid = str(d.get("offerId") or "").strip()
    if not oid:
        return None
    title = clean_title(d.get("title"))
    if not title:
        return None
    pi = d.get("priceInfo") or {}
    price = C._to_float(pi.get("price"))
    if price is None:
        price = C._to_float(pi.get("priceInteger"))
    # 銷量: ★ 2026-11-25 修正 (實測 bug)
    #   afterPrice 在 PC 回應裡是「運費/包郵」說明 (matKey=freight_desc, text="运费3.5元"/"包邮")
    #   → parse_cn_count("运费3.5元") 會抓到「3」→ 銷量被讀成 3 → 好商品被誤殺 (40 筆)
    #   → 口徑改為「bookedCount 為主」(與手機版一致), afterPrice 只在非運費類時才採用
    sales = C._to_float(d.get("bookedCount"))
    _FREIGHT_KEYS = ("freight_desc", "freight", "free_shipping", "free_shipping_desc", "")
    ap = d.get("afterPrice") or {}
    ap_key = str(ap.get("matKey") or "").lower()
    ap_text = str(ap.get("text") or "")
    if sales is None and ap_key not in _FREIGHT_KEYS and ap_text:
        sales = parse_cn_count(ap_text)
    if sales is None:
        for x in (d.get("afterPriceList") or []):
            if not isinstance(x, dict):
                continue
            if str(x.get("matKey") or "").lower() in _FREIGHT_KEYS:
                continue
            sales = parse_cn_count(x.get("text"))
            if sales is not None:
                break
    # 銷量文字 (給 Excel 顯示): 運費/包郵字樣對人沒意義 → 改用「成交N笔」
    sales_text = ap_text
    if (not sales_text) or ap_key in _FREIGHT_KEYS or re.search(r"运费|運費|包邮|包郵", sales_text):
        sales_text = f"成交{sales:.0f}笔" if sales is not None else ""
    # 復購率: afterTags.title=回头率 → afterTags.text=41%; offerRepurchaseRate=14%
    rp = None
    at = d.get("afterTags") or {}
    if str(at.get("title") or "").find("回头率") >= 0:
        rp = C._to_float(at.get("text"))
    if rp is None:
        rp = C._to_float(d.get("offerRepurchaseRate"))
    tp = C._to_float((d.get("shop") or {}).get("tpYear"))
    img = (d.get("offerPicUrl") or "").strip()
    if img.startswith("//"):
        img = "https:" + img
    shop = (d.get("shop") or {}).get("text") or (d.get("shopAddition") or {}).get("text") or ""
    raw_link = d.get("linkUrl") or ""
    if raw_link.startswith("//"):
        raw_link = "https:" + raw_link
    # ★ 一律用標準 offer 連結 (2026-11-10): PC 回應的 linkUrl 是
    #   detailp4p.1688.com/buyer/offerdetail/xxx.html → merge 抽不到 offerId → 圖片對不上
    url = f"https://detail.1688.com/offer/{oid}.html"
    # 額外欄位 (PC 版才有)
    tags = [t.get("text", "") for t in (d.get("tags") or []) if isinstance(t, dict)]
    svc = [s for s in ((d.get("offerTags") or {}).get("serviceTags") or []) if s]
    mid = [m.get("text", "") for m in (d.get("offerMiddle") or []) if isinstance(m, dict)]
    return {
        "offerId": oid,
        "title": title[:120],
        "priceText": f"￥{price:.1f}" if price is not None else "",
        "priceMin": price,
        "priceMax": price,
        "salesText": sales_text,
        "salesNum": sales,
        "repurchaseRate": rp,
        "shopYears": int(tp) if tp is not None else None,
        "mainImage": img,
        "url": url,
        "offerRepurchaseRate": C._to_float(d.get("offerRepurchaseRate")),
        "weight": None,
        "monthSales": None,
        "isNew": False,
        "companyName": shop,
        "pcRawLink": raw_link,
        # PC 專屬
        "pcProvince": d.get("province") or "",
        "pcCity": d.get("city") or "",
        "pcFactory": bool(d.get("factoryInspection")),
        "pcBusiness": bool(d.get("businessInspection")),
        "pcTags": "、".join([t for t in (tags + svc + mid) if t])[:120],
    }


def parse_offers(body: str):
    i, j = body.find("{"), body.rfind("}")
    if i < 0:
        return []
    try:
        obj = json.loads(body[i:j + 1])
    except Exception:
        return []
    try:
        items = obj["data"]["data"]["OFFER"]["items"]
    except Exception:
        return []
    out = []
    for it in items or []:
        d = it.get("data") if isinstance(it, dict) else None
        if isinstance(d, dict):
            out.append(d)
    return out


def main():
    global done_file, kws
    LOG("=" * 64)
    LOG(f"PC 模式採集器  詞 {START}-{END}  tag={TAG}  品類={KWSET}  CDP={CDP_PORT}")
    LOG(f"價帶 {C.PRICE_MIN}-{C.PRICE_MAX}  銷量 {C.MIN_SALES}+"
        f"{'-' + str(C.MAX_SALES) if C.MAX_SALES else ''}  每詞最多 {MAX_PAGES} 頁")
    LOG("=" * 64)

    kw_file = C.KW_FILE
    kws = [l.strip() for l in kw_file.read_text(encoding="utf-8").splitlines() if l.strip()]
    done_file = Path(C.TEMP_DIR) / f"1688_done_v3_{KWSET}.txt"
    done = set(l.strip() for l in done_file.read_text(encoding="utf-8").splitlines()
               if l.strip()) if done_file.exists() else set()
    batch = kws[START - 1:END]

    if not C.acquire_lock(KWSET, tag=TAG):
        LOG(f"[GUARD] 採集鎖被占用 → 中止")
        return 1
    LOG(f"[GUARD] 取得採集鎖 (pid={__import__('os').getpid()})")
    try:
        return _run(batch, done)
    finally:
        C.release_lock()


def _run(batch, done):
    from playwright.sync_api import sync_playwright
    all_items = []
    seen_run = set()
    round_used = 0
    ts_now = C.datetime.now().strftime("%Y-%m-%d %H:%M")

    with sync_playwright() as p:
        br = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
        ctx = br.contexts[0]
        page = None
        for pg in ctx.pages:
            if "s.1688.com" in (pg.url or ""):
                page = pg
                break
        if page is None:
            page = ctx.new_page()
        try:
            page.set_viewport_size({"width": 1400, "height": 900})
        except Exception:
            pass

        bodies = []

        def on_response(resp):
            try:
                if "relationrecommend" in resp.url.lower():
                    bodies.append(resp.text())
            except Exception:
                pass

        page.on("response", on_response)

        for i, kw in enumerate(batch, START):
            if kw in done:
                LOG(f"  [SKIP] 已完成過: {kw}")
                continue
            # ── 守門: 輪次/每日預算 (2026-11-10 補上, PC 通道也要有預算上限) ──
            ok_b, why_b = C.budget_check(KWSET, round_used=round_used)
            if not ok_b:
                LOG(f"[GUARD] {why_b}")
                LOG("[GUARD] 優雅停止 (已收資料都會存檔, 續跑記錄可接著跑)")
                break
            round_used += 1
            C.note_word(KWSET)
            LOG(f"[{i}/{len(kws)}] {kw}")
            bodies.clear()
            got = 0
            for pno in range(1, MAX_PAGES + 1):
                if pno == 1:
                    url = PC_SEARCH.format(enc=gbk_encode(kw))
                else:
                    # 下一頁: 點 .fui-arrow.fui-next (1688-cli 的做法)
                    nxt = page.query_selector(".fui-arrow.fui-next:not(.fui-next-disabled)")
                    if nxt is None:
                        LOG(f"    [分頁] 沒有下一頁 → 停止")
                        break
                    try:
                        nxt.click(timeout=6000)
                    except Exception as e:
                        LOG(f"    [分頁] 點擊失敗: {str(e)[:60]}")
                        break
                    url = None
                bodies.clear()
                try:
                    if url:
                        page.goto(url, wait_until="domcontentloaded", timeout=60000)
                    page.wait_for_timeout(PAGE_SETTLE_MS)
                except Exception as e:
                    LOG(f"    [ERR] 載入失敗: {str(e)[:90]}")
                    break

                # 風控/限流判定 (用頁面層 + 回應層)
                pstate = PS.classify_page_state(page.url or "", page.title() or "",
                                               page.evaluate("() => document.body ? document.body.innerText.slice(0,1500) : ''")
                                               if True else "")
                if pstate[0] in (PS.RISK_CHALLENGE, PS.RATE_LIMITED):
                    LOG(f"    [STOP] {pstate[0]} ({','.join(pstate[1])}) → 停止")
                    C.write_abort_signal(f"PC 模式: {pstate[0]}", kw)
                    break

                items = []
                for b in bodies:
                    items.extend(parse_offers(b))
                if not items:
                    LOG(f"    [WARN] 沒抓到商品 (page{pno})")
                    break
                LOG(f"    page{pno}: {len(items)} 筆 (found 資料)")
                for d in items:
                    it = pc_offer_to_item(d)
                    if not it:
                        continue
                    if it["offerId"] in seen_run:
                        continue
                    ok, why = C.passes_filters(it)
                    if not ok:
                        C.log_dropped(kw, why, it)
                        continue
                    seen_run.add(it["offerId"])
                    it["is_brand"] = False
                    all_items.append(it)
                    got += 1
                if pno < MAX_PAGES:
                    time.sleep(random.uniform(GAP_MIN, GAP_MAX))

            LOG(f"    -> 本詞通過 {got} 筆 (累計 {len(all_items)})")
            if all_items:
                C.save_all_tables(all_items, ts_now)
            time.sleep(random.uniform(GAP_MIN, GAP_MAX))

        # 續跑記錄
        if batch:
            new_done = [k for k in batch if k not in done]
            with done_file.open("a", encoding="utf-8") as f:
                for k in new_done:
                    f.write(k + "\n")
            try:
                bd = Path(C.TEMP_DIR) / f"1688_done_v3_{KWSET}_bydate.txt"
                with bd.open("a", encoding="utf-8") as f:
                    for k in new_done:
                        f.write(f"{C.datetime.now().strftime('%Y-%m-%d')}\t{k}\n")
            except Exception:
                pass

        LOG("")
        LOG(f"[完成] {TAG}: {len(all_items)} 筆, 新 offerId {len(seen_run)}")
        for out in C.save_all_tables(all_items, ts_now) or []:
            LOG(f"    -> {out}")
        br.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
