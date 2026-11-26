# -*- coding: utf-8 -*-
"""final_check_cat.py — 交付前獨立複核 (通用版): 直接讀最終總表抽樣 + 品牌殘留掃描

用法: python final_check_cat.py <cat>
"""
import sys, io, re, random
from pathlib import Path
from openpyxl import load_workbook

sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
from v3_config import CATS, paths

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

CAT = sys.argv[1]
C = CATS[CAT]
F = paths(CAT)["final"]

print(f"檔案: {F}")
print(f"存在: {F.exists()}  大小: {F.stat().st_size:,} bytes")

wb = load_workbook(F)
ws = wb[wb.sheetnames[0]]
print(f"工作表: {wb.sheetnames}")
hdr = [str(c.value) if c.value else "" for c in ws[1]]
print(f"欄位 ({len(hdr)}): {hdr}")

rows = list(ws.iter_rows(min_row=2, values_only=True))
print(f"資料列: {len(rows)}")

ti = hdr.index("商品标题")
pi = hdr.index("价格")
si = hdr.index("销量")
ci = hdr.index("商品链接")

random.seed(20261020)
print("\n=== 隨機抽樣 30 筆 ===")
for r in random.sample(rows, min(30, len(rows))):
    print(f"  {str(r[pi]):>8} {str(r[si]):>10}  {str(r[ti])[:56]}")

print(f"\n空標題: {sum(1 for r in rows if not r[ti])}   空連結: {sum(1 for r in rows if not r[ci])}")

if "序号" in hdr:
    k = hdr.index("序号")
    seq = [r[k] for r in rows]
    print(f"序號 1..{len(rows)} 連續: {seq == list(range(1, len(rows) + 1))}")


def num(v):
    s = str(v or "").replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*万", s)
    if m:
        return float(m.group(1)) * 10000
    m = re.search(r"\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


lo, hi = C["price"]
pr = [num(r[pi]) for r in rows]
sa = [num(r[si]) for r in rows]
pr = [x for x in pr if x is not None]
sa = [x for x in sa if x is not None]
print(f"價格區間 {min(pr)} ~ {max(pr)}  越界({lo}-{hi}): {sum(1 for x in pr if x < lo or x > hi)}")
print(f"銷量區間 {min(sa):.0f} ~ {max(sa):.0f}  越界(<{C['min_sales']}): {sum(1 for x in sa if x < C['min_sales'])}")

# 品牌 / IP 殘留掃描 (國際大牌 + 仿牌 + IP)
BRAND = ["stanley", "膳魔", "膳魔師", "象印", "hydro flask", "hydrojug", "owala",
         "le creuset", "kinto", "hario", "kalita", "chemex", "comandante", "porlex",
         "aeropress", "爱乐压", "愛樂壓", "bialetti", "timemore", "泰摩", "nespresso",
         "smeg", "bodum", "双立人", "wmF".lower(), "coleman", "虎牌", "lock&lock",
         "乐扣", "樂扣", "snow peak", "挪客", "naturehike", "牧高笛", "黑鹿", "伯希和",
         "露得客", "火枫", "火楓", "铠斯", "鎧斯", "拓村", "钛工匠", "鈦工匠",
         "小米", "网易严选", "京东京造", "九阳", "苏泊尔", "美的", "希诺", "思宝",
         "哈尔斯", "富光", "益美", "泰硕", "稻草人", "容山堂", "小罐茶", "三顿半",
         "隅田川", "m stand", "seesaw", "雀巢", "kitty", "三丽鸥", "三麗鷗",
         "史努比", "史迪仔", "布朗熊", "蜘蛛侠", "速度与激情", "故宫文创",
         "疯狂动物城", "雪耀奇境", "delonghi", "德龙", "尤里卡", "eureka",
         "燕印", "tsubame", "世界杯", "史但利", "史单利", "sta-nley", "原始人",
         "探险者", "克米特", "牧高笛", "荣事达", "榮事達", "希乐", "希樂",
         "igloo", "pelican", "rtic", "rubbermaid", "playmate", "decathlon", "迪卡侬",
         "康佳", "万利达", "北鼎", "摩飞", "摩飛", "禾木", "cup"]
hit = []
for r in rows:
    t = str(r[ti]).lower()
    for b in BRAND:
        if b.lower() in t:
            hit.append((b, str(r[ti])[:46]))
            break
print(f"\n品牌/IP 殘留掃描: {len(hit)} 筆")
for b, t in hit[:25]:
    print(f"  [{b}] {t}")
wb.close()
