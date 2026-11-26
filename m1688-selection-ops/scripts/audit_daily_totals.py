# -*- coding: utf-8 -*-
"""audit_daily_totals.py — 統計某品類各批次檔的筆數 (未去重)

用法: python audit_daily_totals.py [品類前綴]
"""
import io
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)

from openpyxl import load_workbook

PREFIX = sys.argv[1] if len(sys.argv) > 1 else "宠物喂食"
EXCEL = Path(r"E:\1shopee\excel")

tot = {}
for f in sorted(EXCEL.glob(f"{PREFIX}-*-bikeparts-*.xlsx")):
    tag = f.name.split("bikeparts-")[1].split("-")[0]
    try:
        wb = load_workbook(f, read_only=True)
        ws = wb.worksheets[0]
        n = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r[1])
        wb.close()
    except Exception:
        continue
    tot[tag] = tot.get(tag, 0) + n

print(f"=== {PREFIX} 各批次筆數 ===")
for t, n in sorted(tot.items(), key=lambda x: -x[1]):
    print(f"  {t:<8} {n:>4}")
print(f"  合計 {sum(tot.values())} 筆 (未去重)")
