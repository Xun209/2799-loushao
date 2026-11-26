# -*- coding: utf-8 -*-
"""merge_v3.py — 某品類全部 v3 批次檔合併 → 去重 → 下載圖片 → 壓縮 → 單一含圖總表
用法: python merge_v3.py <kc|out>

v3 批次檔只存資料 (主圖欄位放的是圖片網址), 圖片在這裡統一並行下載 + 壓縮 + 嵌入。
下載結果快取在 Temp, 重跑不會重複下載。
"""
import sys, io, re, glob, urllib.request
from pathlib import Path
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
from v3_config import CATS, EXCEL, paths, DATE, TMP

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

CAT = sys.argv[1]
C = CATS[CAT]
P = paths(CAT)
prefix = C["prefix"]
IMG_CACHE = TMP / f"v3img_{CAT}"
IMG_FAIL_LOG = []
IMG_CACHE.mkdir(exist_ok=True)

FILES = sorted(glob.glob(str(EXCEL / f"{prefix}-m1688-{C.get('date', DATE)}-bikeparts-*.xlsx")))
FILES = [f for f in FILES if "_test" not in f]
if not FILES:
    print(f"找不到 {prefix} 的批次檔")
    sys.exit(1)

LOG = []
def log(m):
    LOG.append(str(m)); print(m)


def norm_title(t):
    return re.sub(r"[\s\W_]+", "", str(t).lower())


def oid_of(url):
    """從各種 1688 連結形式抽 offerId
       detail.1688.com/offer/xxx.html / detailp4p.1688.com/buyer/offerdetail/xxx.html
       /offer/xxx.html?spm=... / offer_id=xxx"""
    import re as _re
    m = _re.search(r"/offer/(\d+)|offerdetail/(\d+)|offer_id=(\d+)", str(url or ""))
    if m:
        return next((g for g in m.groups() if g), "")
    return ""


def fetch_img(url, oid, retries=3):
    """下載主圖 (帶快取); 回傳 bytes 或 None

    2026-11-10: 加上重試 — CDN 偶發性失敗會造成「缺圖」,
      驗收第一項 (圖片逐列位元組相符) 就過不了。
    """
    if not url:
        return None
    f = IMG_CACHE / f"{oid}.jpg"
    if f.exists() and f.stat().st_size > 1000:
        return f.read_bytes()
    last_err = ""
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 Chrome/120 Safari/537.36",
                "Referer": "https://m.1688.com/",
            })
            with urllib.request.urlopen(req, timeout=25) as r:
                data = r.read()
            if len(data) < 500:
                last_err = f"too small {len(data)}B"
            else:
                f.write_bytes(data)
                return data
        except Exception as e:
            last_err = f"{type(e).__name__}: {str(e)[:60]}"
        time.sleep(1.5 * (attempt + 1))
    if last_err:
        IMG_FAIL_LOG.append(f"{oid}\t{url}\t{last_err}")
    return None


def compress(data):
    """160px / JPEG q72 / optimize — 必須與 verify_v3.py 完全一致"""
    im = PILImage.open(BytesIO(data)).convert("RGB")
    im.thumbnail((160, 160), PILImage.LANCZOS)
    b = BytesIO()
    im.save(b, format="JPEG", quality=72, optimize=True)
    return b.getvalue()


header = widths = None
pool = []
for f in FILES:
    p = Path(f)
    wb = load_workbook(p)
    ws = wb[wb.sheetnames[0]]
    hdr = [str(c.value) if c.value else "" for c in ws[1]]
    if header is None:
        header = hdr
        widths = {k: v.width for k, v in ws.column_dimensions.items() if v.width}
    n = 0
    for r in ws.iter_rows(min_row=2):
        d = dict(zip(hdr, [c.value for c in r]))
        if not d.get("商品标题"):
            continue
        n += 1
        u = str(d.get("商品链接") or "")
        oid = oid_of(u) or str(r[0].row)
        pool.append({
            "src": p.name, "d": d, "oid": oid,
            "title": str(d["商品标题"]).strip(),
            "url": u,
            "img": str(d.get("主图(嵌图)") or "").strip(),
            "nt": norm_title(d["商品标题"]),
        })
    log(f"[FILE] {p.name}: {n} 條")
    wb.close()

log(f"\n[合計] {len(pool)} 筆")

seen_url, seen_nt, final, dup_url, dup_nt = set(), set(), [], 0, 0
by_nt = {}
for rec in pool:
    dup = False
    if rec["url"] and rec["url"] in seen_url:
        dup = True; dup_url += 1
    elif rec["nt"] in seen_nt:
        dup = True; dup_nt += 1
    if dup:
        # 重複商品: 若已保留的那筆沒有圖片網址, 就用有圖片的這筆補上
        prev = by_nt.get(rec["nt"])
        if prev is not None and not prev.get("img") and rec.get("img"):
            prev["img"] = rec["img"]
            prev["d"]["主图(嵌图)"] = rec["img"]      # 儲存格內容也要補, 否則欄位是空的
        continue
    if rec["url"]:
        seen_url.add(rec["url"])
    seen_nt.add(rec["nt"])
    by_nt[rec["nt"]] = rec
    final.append(rec)
log(f"[去重] URL 重複 {dup_url}, 標題重複 {dup_nt} -> 最終 {len(final)} 筆")

log("[IMG] 並行下載主圖中 ...")
def dl(rec):
    return rec["oid"], fetch_img(rec["img"], rec["oid"])

img_map = {}
with ThreadPoolExecutor(max_workers=8) as ex:
    for oid, data in ex.map(dl, final):
        if data:
            img_map[oid] = data
n_ok = len(img_map)
log(f"[IMG] 下載成功 {n_ok} / {len(final)}")

wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = C["sheet"]
for c, h in enumerate(header, start=1):
    ws_out.cell(row=1, column=c, value=h)
ws_out.row_dimensions[1].height = 32
if widths:
    for k, w in widths.items():
        ws_out.column_dimensions[k].width = w

c_seq = header.index("序号") + 1

for i, rec in enumerate(final, start=2):
    for c, h in enumerate(header, start=1):
        v = rec["d"].get(h)
        if v is None:
            continue
        if c == c_seq:
            v = i - 1
        ws_out.cell(row=i, column=c, value=v)
    ws_out.row_dimensions[i].height = 120

wb_out.save(str(P["merge"]))
log(f"\n[SAVE] {P['merge']}")
log(f"  [IMG] 下載失敗 {len(IMG_FAIL_LOG)} 筆")
log(f"  圖片快取 {n_ok} / {len(final)} (原始 bytes, 複審階段才壓縮嵌入)")
log(f"  檔案大小 {P['merge'].stat().st_size:,} bytes")
P["log"].write_text("\n".join(LOG), encoding="utf-8")
