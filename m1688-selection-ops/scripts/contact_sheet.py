# -*- coding: utf-8 -*-
"""contact_sheet.py — 從最終總表抽圖 + 標題做對照表, 視覺驗證主圖與標題是否對位
用法: python contact_sheet.py [cat] [起始列] [張數]
"""
import sys, io
from pathlib import Path
from io import BytesIO
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
from v3_config import CATS, paths
from xlsx_imgfix import images_by_row

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)

CAT = sys.argv[1] if len(sys.argv) > 1 else "wall"
START = int(sys.argv[2]) if len(sys.argv) > 2 else 1
N = int(sys.argv[3]) if len(sys.argv) > 3 else 24
C = CATS[CAT]
P = paths(CAT)
OUT = Path(r"E:\1shopee\輸出") / f"{CAT}-contact-sheet-{START}.png"

wb = load_workbook(P["final"])
ws = wb[wb.sheetnames[0]]
hdr = [str(c.value) if c.value else "" for c in ws[1]]
ti, pi, si, ci = hdr.index("商品标题"), hdr.index("价格"), hdr.index("销量"), hdr.index("商品链接")

imgs = images_by_row(P["final"])
print(f"[{CAT}] 最終總表 {ws.max_row - 1} 筆, 抽圖到 {len(imgs)} 列有圖")

CELL_W, CELL_H = 300, 330
COLS = 6
rows_n = (min(N, ws.max_row - 1) + COLS - 1) // COLS
sheet = Image.new("RGB", (CELL_W * COLS, CELL_H * rows_n), (250, 250, 250))
draw = ImageDraw.Draw(sheet)
try:
    font = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 15)
    font_s = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 13)
except Exception:
    font = ImageFont.load_default()
    font_s = font

picked = 0
for i in range(START, min(START + N, ws.max_row)):
    row = i + 1          # Excel 列號 = 序號 + 1 (第1列是標題列)
    idx = (i - START) % COLS
    ry = (i - START) // COLS
    x0, y0 = idx * CELL_W, ry * CELL_H
    title = str(ws.cell(row=row, column=ti + 1).value or "")
    price = str(ws.cell(row=row, column=pi + 1).value or "")
    sales = str(ws.cell(row=row, column=si + 1).value or "")
    link = str(ws.cell(row=row, column=ci + 1).value or "")

    got = imgs.get(row) or []
    if got:
        try:
            im = Image.open(BytesIO(got[0][0])).convert("RGB")
            im.thumbnail((CELL_W - 20, 150))
            sheet.paste(im, (x0 + (CELL_W - im.width) // 2, y0 + 4))
        except Exception:
            draw.rectangle([x0 + 10, y0 + 10, x0 + CELL_W - 10, y0 + 150], outline=(200, 0, 0), width=2)
    else:
        draw.rectangle([x0 + 10, y0 + 10, x0 + CELL_W - 10, y0 + 150], outline=(255, 0, 0), width=3)
        draw.text((x0 + 30, y0 + 70), "NO IMAGE", fill=(255, 0, 0), font=font)
    draw.text((x0 + 8, y0 + 158), f"#{i}  {price}  {sales}", fill=(0, 0, 0), font=font_s)
    # 標題換行 (每行 20 字)
    t = title
    for k in range(0, min(len(t), 60), 20):
        draw.text((x0 + 8, y0 + 178 + (k // 20) * 17), t[k:k + 20], fill=(40, 40, 40), font=font_s)
    oid = link.split("/offer/")[1].split(".html")[0] if "/offer/" in link else "?"
    draw.text((x0 + 8, y0 + CELL_H - 20), oid, fill=(120, 120, 120), font=font_s)
    picked += 1

sheet.save(OUT)
print(f"已輸出 {picked} 筆對照圖 -> {OUT}")
print(f"尺寸 {sheet.size}")
wb.close()
