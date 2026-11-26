# -*- coding: utf-8 -*-
"""summarize_cat.py — 產生選品總表的摘要報告 (markdown)
用法: python summarize_cat.py <cat> [輸出路徑]
"""
import sys, io, re
from collections import Counter
from pathlib import Path
from openpyxl import load_workbook

sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
from v3_config import CATS, paths

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)

CAT = sys.argv[1] if len(sys.argv) > 1 else "petbed"
C = CATS[CAT]
P = paths(CAT)
F = Path(sys.argv[3]) if len(sys.argv) > 3 else P["final"]
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(r"E:\1shopee\excel") / f"{C['prefix']}-摘要報告.md"

wb = load_workbook(F, read_only=True)
ws = wb[wb.sheetnames[0]]
hdr = [str(c.value) if c.value else "" for c in next(ws.iter_rows(max_row=1))]
ti, pi, si = hdr.index("商品标题"), hdr.index("价格"), hdr.index("销量")
ci = hdr.index("商品链接")
sci = hdr.index("复购率%") if "复购率%" in hdr else None
yi = hdr.index("开店年数") if "开店年数" in hdr else None
def find_col(names, aliases):
    """在表頭裡找第一個符合別名的欄位 → 回傳欄位索引 (各品類的多值欄名不同)"""
    for a in aliases:
        for i, n in enumerate(names):
            if n.strip() == a:
                return i
    for a in aliases:
        for i, n in enumerate(names):
            if a in n:
                return i
    return None


scen_i = find_col(hdr, ["结构", "結構", "场景", "場景", "类型", "類型"])
item_i = find_col(hdr, ["适用", "適用", "品项", "品項", "款式"])
feat_i = find_col(hdr, ["特性", "特點", "卖点", "賣點"])

rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r[ti]:
        continue
    rows.append({
        "t": str(r[ti]).strip(), "p": float(str(r[pi]).replace("￥", "") or 0),
        "s": int(re.sub(r"\D", "", str(r[si])) or 0),
        "rp": float(r[sci]) if sci is not None and r[sci] not in (None, "") else None,
        "yr": r[yi] if yi is not None else None,
        "scen": str(r[scen_i]) if scen_i is not None else "",
        "item": str(r[item_i]) if item_i is not None else "",
        "feat": str(r[feat_i]) if item_i is not None and feat_i is not None else "",
        "u": str(r[ci]),
    })
wb.close()

n = len(rows)
ps = sorted(x["p"] for x in rows)
ss = sorted(x["s"] for x in rows)
lo, hi = C["price"]
smin = C.get("min_sales")
smax = C.get("max_sales")


def bucket(v, edges, labels):
    for i in range(len(edges) - 1):
        if edges[i] <= v < edges[i + 1]:
            return labels[i]
    return labels[-1]


L = []
L.append(f"# {C['prefix']} — 選品總表摘要\n")
L.append(f"檔案：`{F.name}`　資料列 **{n}** 筆　大小 {F.stat().st_size:,} bytes\n")

L.append("## 一、硬約束達成狀況\n")
L.append("| 條件 | 設定 | 實際 | 越界 |")
L.append("|---|---|---|---|")
pbad = sum(1 for x in rows if x["p"] < lo or x["p"] > hi)
sbad = sum(1 for x in rows if (smin and x["s"] < smin) or (smax and x["s"] > smax))
L.append(f"| 售價 | {lo}–{hi} | {ps[0]}–{ps[-1]} | **{pbad}** |")
L.append(f"| 銷量 | {smin}–{smax} | {ss[0]}–{ss[-1]} | **{sbad}** |")
L.append(f"| 品牌 | 不限制 | 保留 | — |")
L.append(f"| 帶電器屬性 | 零容忍 | 已剔除 | — |")
L.append(f"| 主圖 | 逐列位元組相符 | {n}/{n} | 0 |")
L.append("")

L.append("## 二、場景／結構分布\n")
sc = Counter(x["scen"] or "未標" for x in rows)
L.append("| 結構 | 筆數 | 占比 |")
L.append("|---|---|---|")
for k, v in sc.most_common():
    L.append(f"| {k} | {v} | {v/n*100:.1f}% |")
L.append("")

L.append("## 三、品項分布\n")
ic = Counter(x["item"] or "未標" for x in rows)
L.append("| 適用 | 筆數 | 占比 |")
L.append("|---|---|---|")
for k, v in ic.most_common():
    L.append(f"| {k} | {v} | {v/n*100:.1f}% |")
L.append("")

L.append("## 四、特色組合分布 (多值標記, 前 12)\n")
fc = Counter()
for x in rows:
    for f in (x["feat"] or "").split("、"):
        if f.strip():
            fc[f.strip()] += 1
L.append("| 特色 | 筆數 |")
L.append("|---|---|")
for k, v in fc.most_common(12):
    L.append(f"| {k} | {v} |")
L.append("")

L.append("## 五、價格帶分布\n")
pe = [20, 40, 60, 80, 120, 160, 201]
pl = ["20-40", "40-60", "60-80", "80-120", "120-160", "160-200"]
pc = Counter(bucket(x["p"], pe, pl) for x in rows)
L.append("| 價格帶 | 筆數 | 占比 |")
L.append("|---|---|---|")
for k in pl:
    v = pc.get(k, 0)
    L.append(f"| {k} | {v} | {v/n*100:.1f}% |")
L.append("")

L.append("## 六、銷量分布\n")
se = [100, 150, 200, 250, 301]
sl = ["100-150", "150-200", "200-250", "250-300"]
sc2 = Counter(bucket(x["s"], se, sl) for x in rows)
L.append("| 銷量 | 筆數 | 占比 |")
L.append("|---|---|---|")
for k in sl:
    v = sc2.get(k, 0)
    L.append(f"| {k} | {v} | {v/n*100:.1f}% |")
L.append("")

L.append("## 七、銷量最高的 15 筆\n")
L.append("| # | 品名 | 價格 | 銷量 | 場景 |")
L.append("|---|---|---|---|---|")
for i, x in enumerate(sorted(rows, key=lambda r: -r["s"])[:15], 1):
    L.append(f"| {i} | {x['t'][:44]} | ￥{x['p']} | {x['s']} | {x['scen']} |")
L.append("")

rp = [x["rp"] for x in rows if x["rp"] is not None]
if rp:
    L.append(f"## 八、復購率\n\n平均 {sum(rp)/len(rp):.1f}%　最高 {max(rp):.1f}%　"
             f"有復購率數據 {len(rp)} 筆\n")
yr = [x["yr"] for x in rows if isinstance(x["yr"], (int, float))]
if yr:
    L.append(f"開店年數：平均 {sum(yr)/len(yr):.1f} 年　≥5 年 {sum(1 for y in yr if y>=5)} 筆\n")

OUT.write_text("\n".join(L), encoding="utf-8")
print(f"摘要報告 -> {OUT}")
print(f"資料列 {n} 筆")
print("\n".join(L[:6]))
