# -*- coding: utf-8 -*-
"""verify_v3.py — 最終總表逐列驗證
  1. 圖片 = 合併總表對應列圖片 (同樣壓縮參數比對)
  2. 順序 / 序號連續 / 價格範圍 / 銷量範圍 / 列高
用法: python verify_v3.py <kc|out>
"""
import sys, io, re
from pathlib import Path
from io import BytesIO
from openpyxl import load_workbook
from PIL import Image as PILImage

sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
from xlsx_imgfix import images_by_row
from v3_config import CATS, paths, oid_of, raw_img, EXCEL
from v3_rules import hit_reason, parse_sales

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

CAT = sys.argv[1]
C = CATS[CAT]
P = paths(CAT)


def norm_title(t):
    return re.sub(r"[\s\W_]+", "", str(t).lower())


# 跨表去重 (必須與 finalize_v3.py 同一套邏輯, 否則期望值會對不上)
DEDUP_AGAINST = C.get("dedup_against") or []
dup_oids, dup_titles = set(), set()
for pat in DEDUP_AGAINST:
    for f in EXCEL.glob(pat):
        if f.name.startswith("~$"):
            continue
        try:
            wb2 = load_workbook(f, read_only=True)
            ws2 = wb2[wb2.sheetnames[0]]
            h2 = [str(c.value) if c.value else "" for c in next(ws2.iter_rows(max_row=1))]
            t2 = h2.index("商品标题")
            l2 = h2.index("商品链接")
            for r2 in ws2.iter_rows(min_row=2, values_only=True):
                if not r2[t2]:
                    continue
                u = str(r2[l2] or "")
                if "/offer/" in u:
                    dup_oids.add(u.split("/offer/")[1].split(".html")[0])
                dup_titles.add(norm_title(r2[t2]))
            wb2.close()
        except Exception:
            pass


def shrink(data):
    im = PILImage.open(BytesIO(data)).convert("RGB")
    im.thumbnail((160, 160), PILImage.LANCZOS)
    buf = BytesIO()
    im.save(buf, format="JPEG", quality=72, optimize=True)
    return buf.getvalue()


wb = load_workbook(P["merge"]); ws = wb[wb.sheetnames[0]]
h = [str(c.value) if c.value else "" for c in ws[1]]
ct = h.index("商品标题")
cs = h.index("销量")
cl = h.index("商品链接")
keep = []
for r in ws.iter_rows(min_row=2):
    if not r[ct].value:
        continue
    t = str(r[ct].value).strip()
    sv = parse_sales(r[cs].value)
    if sv is not None and sv < C["min_sales"]:
        continue
    # 銷量上限 (可選, 2026-11-04 新增)
    max_sales = C.get("max_sales")
    if max_sales is not None and sv is not None and sv > max_sales:
        continue
    if dup_oids or dup_titles:
        o = oid_of(r[cl].value)
        if (o and o in dup_oids) or norm_title(t) in dup_titles:
            continue
    if hit_reason(C, t):
        continue
    keep.append((oid_of(r[cl].value), t))
wb.close()
print("預期保留:", len(keep))

wb = load_workbook(P["final"]); ws = wb[wb.sheetnames[0]]
hh = [str(c.value) if c.value else "" for c in ws[1]]
rows = list(ws.iter_rows(min_row=2, values_only=True))
out_imgs = images_by_row(P["final"])
print("實際輸出:", len(rows))

ok = bad = 0
badex = []
for i, (oid, t) in enumerate(keep, start=2):
    if i - 2 >= len(rows):
        bad += 1
        badex.append(f"輸出列數不足 row{i}")
        continue
    got = str(rows[i - 2][hh.index("商品标题")])
    if got != t:
        bad += 1
        if len(badex) < 6:
            badex.append(f"row{i} 順序錯: 期望[{t[:24]}] 實際[{got[:24]}]")
        continue
    raw = raw_img(CAT, oid)
    o = out_imgs.get(i) or []
    if not raw or not o:
        bad += 1
        if len(badex) < 6:
            badex.append(f"row{i} 缺圖: {t[:30]}")
        continue
    if shrink(raw) == o[0][0]:
        ok += 1
    else:
        bad += 1
        if len(badex) < 6:
            badex.append(f"row{i} 圖片不符: {t[:34]}")

print(f"\n圖片逐列比對 (同樣壓縮後): 正確 {ok} / 錯 {bad}")
for x in badex:
    print("  ", x)

seq = [rows[i][hh.index("序号")] for i in range(len(rows))]
exp = list(range(1, len(rows) + 1))
print(f"序號連續 1..{len(rows)}: {'OK' if seq == exp else 'FAIL'}")

def num(v):
    m = re.search(r"(\d+(?:\.\d+)?)\s*万", str(v or ""))
    if m:
        return float(m.group(1)) * 10000
    m = re.search(r"\d+(?:\.\d+)?", str(v or "").replace(",", ""))
    return float(m.group()) if m else None

lo, hi = C["price"]
pr = [num(rows[i][hh.index("价格")]) for i in range(len(rows))]
pr = [x for x in pr if x is not None]
sa = [num(rows[i][hh.index("销量")]) for i in range(len(rows))]
sa = [x for x in sa if x is not None]
print(f"價格 {min(pr)}-{max(pr)} (應 {lo}-{hi})  越界 {sum(1 for x in pr if x < lo or x > hi)} 筆")
MAX_SALES = C.get("max_sales")
if MAX_SALES is not None:
    print(f"銷量 {min(sa)}-{max(sa)} (應 {C['min_sales']}-{MAX_SALES})  "
          f"越界 {sum(1 for x in sa if x < C['min_sales'] or x > MAX_SALES)} 筆")
else:
    print(f"銷量 {min(sa)}-{max(sa)} (應 >={C['min_sales']})  越界 {sum(1 for x in sa if x < C['min_sales'])} 筆")
hs = set(ws.row_dimensions[i].height for i in range(2, len(rows) + 2))
print(f"資料列高集合: {hs} (應全為 120)")
print(f"標題列高: {ws.row_dimensions[1].height}")
# 空標題 / 空連結檢查
bad_link = sum(1 for i in range(len(rows)) if not rows[i][hh.index("商品链接")])
print(f"空商品連結: {bad_link} 筆")
wb.close()
