# -*- coding: utf-8 -*-
"""finalize_v3.py — 最終複審: 刪品牌/非本類 → 重編號 → 圖片壓縮 → 最終總表
圖片 160px / JPEG quality 72 (thumbnail 保持比例); 列高 標題 32pt、資料列 120pt
用法: python finalize_v3.py <kc|out>
"""
import sys, io, re
from pathlib import Path
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
from v3_config import CATS, paths, TMP, oid_of, raw_img, EXCEL
from v3_rules import hit_reason, parse_sales

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

CAT = sys.argv[1]
C = CATS[CAT]
P = paths(CAT)

LOG = []
def log(m):
    LOG.append(str(m)); print(m)


def _norm_title(t):
    return re.sub(r"[\s\W_]+", "", str(t).lower())


# ── 跨表去重: 與指定「已交付總表」重複的品項直接剔除 ────────────────────
# 用途: 同一品類做第二彈時, 採集器只擋 offerId, 但同一款商品在 1688 可能
#       有不同賣家的不同連結 (offerId 不同、標題相同) → 這裡用「標準化標題」再擋一次。
# 品類設定: "dedup_against": ["<檔名 glob>"]
DEDUP_AGAINST = C.get("dedup_against") or []
dup_oids, dup_titles = set(), set()
if DEDUP_AGAINST:
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
                    dup_titles.add(_norm_title(r2[t2]))
                wb2.close()
            except Exception as e:
                log(f"[WARN] 讀去重來源失敗 {f.name}: {e}")
    log(f"[DEDUP] 對照來源: {len(dup_oids)} 個 offerId / {len(dup_titles)} 個標準化標題")


wb = load_workbook(P["merge"])
ws = wb[wb.sheetnames[0]]
hdr = [str(c.value) if c.value else "" for c in ws[1]]
widths = {k: v.width for k, v in ws.column_dimensions.items() if v.width}

kept, dropped = [], {}
for r in ws.iter_rows(min_row=2):
    d = dict(zip(hdr, [c.value for c in r]))
    if not d.get("商品标题"):
        continue
    title = str(d["商品标题"]).strip()
    sv = parse_sales(d.get("销量"))
    if sv is not None and sv < C["min_sales"]:
        dropped[f"銷量不足(<{C['min_sales']})"] = dropped.get(f"銷量不足(<{C['min_sales']})", 0) + 1
        continue
    # 銷量上限 (可選, 2026-11-04 新增; 沒有此設定的品類維持原行為)
    max_sales = C.get("max_sales")
    if max_sales is not None and sv is not None and sv > max_sales:
        dropped[f"銷量超上限(>{max_sales})"] = dropped.get(f"銷量超上限(>{max_sales})", 0) + 1
        continue
    # 跨表去重 (與已交付總表重複)
    if dup_oids or dup_titles:
        oid = oid_of(d.get("商品链接"))
        if (oid and oid in dup_oids) or _norm_title(title) in dup_titles:
            dropped["與已交付總表重複"] = dropped.get("與已交付總表重複", 0) + 1
            continue
    reason = hit_reason(C, title)
    if reason:
        dropped[reason] = dropped.get(reason, 0) + 1
        continue
    kept.append({"d": d, "title": title, "oid": oid_of(d.get("商品链接"))})
wb.close()

log("=== 刪除統計 ===")
tot = 0
for k, v in sorted(dropped.items(), key=lambda x: -x[1]):
    log(f"  {v:4d}  {k}")
    tot += v
log(f"  合計刪除 {tot} 筆")
log(f"  保留 {len(kept)} 筆")

wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = C["sheet"]
for c, h in enumerate(hdr, start=1):
    ws_out.cell(row=1, column=c, value=h)
ws_out.row_dimensions[1].height = 32
for k, w in widths.items():
    ws_out.column_dimensions[k].width = w

c_seq = next(i for i, h in enumerate(hdr) if "序号" in h) + 1
c_img = next(i for i, h in enumerate(hdr) if "主图" in h) + 1
img_col = chr(ord("A") + c_img - 1)

ok = fail = 0
for i, rec in enumerate(kept, start=2):
    for c, h in enumerate(hdr, start=1):
        v = rec["d"].get(h)
        if v is None:
            continue
        if c == c_seq:
            v = i - 1
        ws_out.cell(row=i, column=c, value=v)
    ws_out.row_dimensions[i].height = 120
    data = raw_img(CAT, rec["oid"])
    if data:
        try:
            im = PILImage.open(BytesIO(data)).convert("RGB")
            # 用 thumbnail 保持原始寬高比 (不可用 resize, 會變形)
            im.thumbnail((160, 160), PILImage.LANCZOS)
            b = BytesIO()
            im.save(b, format="JPEG", quality=72, optimize=True)
            p = TMP / f"_v3F_{CAT}_{i}.jpg"
            p.write_bytes(b.getvalue())
            x = XLImage(str(p))
            x.width = 100; x.height = 100
            ws_out.add_image(x, f"{img_col}{i}")
            ok += 1
        except Exception as e:
            fail += 1; log(f"  [IMG FAIL] row{i}: {e}")
    else:
        fail += 1

# ── 額外欄位 (品類設定的 extra_cols) ───────────────────────────────────
# 加在**最後**一欄: 前面的欄位與圖片錨點位置都不動, verify 的欄位查找也不受影響
# 格式: (欄位名, [(正則, 值), ...], 預設值) — 依序比對, 第一個命中就採用
EXTRA = C.get("extra_cols") or []
for ec_i, (ec_name, ec_groups, ec_default) in enumerate(EXTRA):
    ec_col = len(hdr) + 1 + ec_i
    ws_out.cell(row=1, column=ec_col, value=ec_name)
    ws_out.column_dimensions[chr(ord("A") + ec_col - 1)].width = 14
    cnt = {}
    # 多值模式 (2026-11-04 新增): 群組中有任一項標 "multi" → 收集全部命中並用「、」串接
    # 沒有標記的既有品類維持「第一個命中就停」的原行為
    is_multi = any(len(g) > 2 and g[2] == "multi" for g in ec_groups)
    for i, rec in enumerate(kept, start=2):
        hits = []
        for g in ec_groups:
            pat, name = g[0], g[1]
            if re.search(pat, rec["title"]):
                hits.append(name)
                if not is_multi:
                    break
        if not hits:
            val = ec_default
        elif is_multi:
            val = "、".join(dict.fromkeys(hits))
        else:
            val = hits[0]
        ws_out.cell(row=i, column=ec_col, value=val)
        cnt[val] = cnt.get(val, 0) + 1
    log(f"[EXTRA] {ec_name}: {cnt}")

wb_out.save(str(P["final"]))
log(f"\n[SAVE] {P['final']}")
log(f"  圖片 {ok}, 缺圖 {fail}")
log(f"  檔案大小 {P['final'].stat().st_size:,} bytes")
P["log"].write_text("\n".join(LOG), encoding="utf-8")
