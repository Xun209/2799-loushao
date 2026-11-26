"""
collect_m1688_v3_api.py — 1688 選品採集 v3 (直調 mtop API 版, 2026-10-20)

重大改版: 不再「載入搜尋頁 + 捲動 + 爬 DOM」, 改為直接呼叫 1688 手機版內部搜尋接口
    h5api.m.1688.com/h5/mtop.1688.wap.seo.offer.get/1.0/
  - 簽名: sign = md5(_m_h5_tk前段 & t & appKey & data)   (appKey=12574478)
  - 分頁: data.beginPage / data.pageSize  (實測 pageSize 最大 100, 200 會被砍到 60)
  - 請求在「頁面 context 內 fetch」→ 沿用 Chromium 真正的 TLS 指紋 + cookie + CORS
  - 優點: 不觸發 document 層的搜索配額、欄位是結構化 JSON、單詞 3-5 秒完成

用法: python collect_m1688_v3_api.py <start> <end> <tag> <詞表檔號>
  詞表檔號: out=戶外便攜休息用品 / kc=兒童滑板車
  例: python collect_m1688_v3_api.py 1 12 o6 out
"""
import hashlib
import io
import json
import os
import random
import sys
import re
import tempfile
import threading
import time
import urllib.request
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace", line_buffering=True)
        # 保留參照: 這兩個 wrapper 若被 GC 會關掉底層 buffer,
        # 導致「import 本模組」的程式出現 "I/O operation on closed file"
        _STDOUT_WRAPPER = sys.stdout
        _STDERR_WRAPPER = sys.stderr
    except Exception:
        pass

TEMP_DIR = Path(tempfile.gettempdir())
SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_DIR = Path("E:/1shopee/excel")
CDP_URL = os.environ.get("M1688_CDP", "http://127.0.0.1:9223")
ABORT_FILE = TEMP_DIR / "1688_abort_signal.txt"
EXISTING_OIDS_FILE = EXCEL_DIR / "_existing_offerids_2026-08-27.txt"

# ── 節奏/頁數可用環境變數微調 (預設值與原本完全相同) ────────────────────
# 2026-11-04: 帳號被 RGV587 限流後, 可用 M1688_MAX_PAGES / M1688_KW_GAP_* 等
# 把請求速率壓低, 換取能多跑幾個詞而不觸發滑塊攔截。
def _envf(name, default):
    try:
        return float(os.environ.get(name, default))
    except Exception:
        return float(default)


def _envi(name, default):
    try:
        return int(float(os.environ.get(name, default)))
    except Exception:
        return int(default)

# ── mtop 接口參數 (2026-10-20 實測) ─────────────────────────────────
MTOP_API = "mtop.1688.wap.seo.offer.get"
MTOP_APP_KEY = "12574478"          # 簽名用的 appKey
DATA_APP_KEY = "1772295d201e06bd6d56a6ed5cfd252f"   # data 內的 appKey
PAGE_SIZE = 60                     # 實測 60/100 都可, 200 被砍成 60
MAX_PAGES_PER_KW = _envi("M1688_MAX_PAGES", 3)   # 單詞最多翻 3 頁 = 180 筆
TARGET_PER_KW = _envi("M1688_TARGET", 30)         # 單詞「通過篩選」達 30 筆就提早停, 省請求
# ── 排序 + 價格帶 (2026-10-20 實測, 命中率的關鍵) ────────────────────
# sortType:
#   "pop" = 綜合排序 -> 偏新品/低銷量, 命中率差 (4/60)
#   任何其他值 -> 後端落到「預設排序」(銷量傾向), 命中率好 (14~41/60)
#   已驗證 price_asc / zzz_not_a_sort / SALE_DESC_XYZ 三者結果完全相同
# priceStart/priceEnd: 後端認得, 會把結果壓進價格帶 (價 min 80 / max 339)
SORT_TYPE = "price_asc"

START = int(sys.argv[1]) if len(sys.argv) > 1 else 1
END = int(sys.argv[2]) if len(sys.argv) > 2 else 50
TAG = sys.argv[3] if len(sys.argv) > 3 else "v1"
_KWSET = sys.argv[4] if len(sys.argv) > 4 else "out"
DATE = "2026-10-20"
LOG_FILE = TEMP_DIR / f"v3api_{_KWSET}_{TAG}_run.log"

# ── 品類設定 ────────────────────────────────────────────────────────
_HEAVY_COMMON = ["石磨", "实心球", "實心球", "沙袋", "负重背心", "負重背心"]

CONFIG = {
    "out": {
        "kwfile": "keywords_outdoor.txt",
        "prefix": "户外便携休息用品",
        "price": (80.0, 350.0),
        "min_sales": 50,
        "brand": [
            "Coleman", "COLEMAN", "科勒曼", "科曼",
            "Naturehike", "NATUREHIKE", "挪客", "Naturehike挪客",
            "Snow Peak", "SNOWPEAK", "雪峰",
            "Mobi Garden", "MOBIGARDEN", "牧高笛",
            "KingCamp", "KINGCAMP",
            "Blackdeer", "BLACKDEER", "黑鹿",
            "Helinox", "HELINOX", "黑诺",
            "Therm-a-Rest", "ThermARest", "Therm-a-rest",
            "Sea to Summit", "SeaToSummit",
            "Big Agnes", "REI Co-op",
            "Klymit", "KLYMIT", "Intex", "INTEX",
            "Decathlon", "DECATHLON", "迪卡侬",
            "3F UL Gear", "Ozark", "OZARK",
            "Marmot", "MARMOT", "山之客", "The North Face", "TNF",
            "Jack Wolfskin", "狼爪", "CAMEL", "骆驼", "CRAGHOPPERS",
            "CAMPFIRE", "Campfire", "Chill Outdoor", "閒風藍", "闲风蓝",
            "露營先生", "露营先生", "TILLAK", "SWS WELLHIKE", "WELLHIKE",
            "DOD", "露得客", "戶外家",
        ],
        "exclude": [
            "康复", "康復", "理疗", "理療", "治疗", "治療", "医用", "醫用", "保健",
            "颈椎", "頸椎", "腰椎", "关节炎", "關節炎", "术后", "術後",
            "血压", "血壓", "血糖", "病人", "残障", "殘障", "轮椅", "輪椅",
            "二手", "中古", "回收", "改装店", "維修店", "维修店", "闲置", "閒置",
            "贴牌", "貼牌", "代工", "定制logo", "定製logo", "贴标", "貼標",
            "宠物", "寵物", "狗", "猫", "貓",
            "衣服", "服裝", "服饰", "服飾", "短袖", "长袖", "長袖",
            "T恤", "t恤", "T恤衫", "裤子", "褲子", "运动裤", "運動裤",
            "运动服", "運動服", "训练服", "訓練服", "健身服", "瑜伽服", "瑜珈服",
            "运动内衣", "運動內衣", "文胸", "内衣", "內衣",
            "压缩裤", "壓縮裤", "紧身衣", "緊身衣", "legging", "泳衣", "泳镜", "泳帽",
            "蛋白粉", "氮泵", "肌酸", "維生素", "维生素", "保健品", "减肥药", "減肥藥",
            "婴儿床", "嬰兒床", "游戏围栏", "遊戲圍欄",
        ],
        "heavy": _HEAVY_COMMON,
        # 同品類舊 DOM 批次檔: 這些 offerId 要保留 (跨任務去重池會擋, 所以開白名單)
        "keep_files": [
            "户外便携休息用品-m1688-2026-10-19-bikeparts-o1-表01.xlsx",
            "户外便携休息用品-m1688-2026-10-19-bikeparts-o2-表01.xlsx",
            # o3 是舊程式產的 (主圖欄位沒存網址), 重採一次才能有圖
            "户外便携休息用品-m1688-2026-10-20-bikeparts-o3-表01.xlsx",
        ],
    },
    "kc": {
        "kwfile": "keywords_kidscooter.txt",
        "prefix": "儿童滑板车",
        "price": (80.0, 350.0),
        "min_sales": 100,
        "brand": [
            "小米", "Xiaomi", "XIAOMI", "米家", "MIJIA", "Mijia",
            "九号", "九號", "Segway", "SEGWAY", "Ninebot", "NINEBOT", "赛格威", "賽格威",
            "Razor", "RAZOR", "Micro", "MICRO", "Mini Micro", "Lil' Rider", "LilRider",
            "LIL RIDER", "Alien Rides", "Dynacraft", "DYNACRAFT", "Schwinn", "SCHWINN",
            "Raleigh", "Tony Hawk", "Yvolution", "YVOLUTION",
            "Fisher-Price", "FisherPrice", "费雪", "費雪",
            "Chicco", "CHICCO", "智高", "Disney", "DISNEY", "迪士尼",
            "Decathlon", "DECATHLON", "迪卡侬",
            "Bilibo", "BILIBO", "贝乐宝", "貝樂寶",
            "小爵士", "Goodbaby", "GoodBaby", "好孩子",
        ],
        "exclude": [
            "康复", "康復", "理疗", "理療", "治疗", "治療", "医用", "醫用", "保健",
            "颈椎", "頸椎", "腰椎", "关节炎", "關節炎", "术后", "術後",
            "血压", "血壓", "血糖", "病人", "残障", "殘障",
            "电动滑板车", "電動滑板車", "电动滑板", "電動滑板",
            "续航", "續航", "锂电池", "鋰電池", "电池包", "電池包",
            "二手", "中古", "回收", "改装店", "維修店", "维修店", "闲置", "閒置",
            "贴牌", "貼牌", "代工", "定制logo", "定製logo", "贴标", "貼標",
            "宠物", "寵物", "狗", "猫", "貓",
            "衣服", "服裝", "服饰", "服飾", "短袖", "长袖", "長袖",
            "T恤", "t恤", "T恤衫", "裤子", "褲子", "运动裤", "運動裤",
            "运动服", "運動服", "训练服", "訓練服", "健身服", "瑜伽服", "瑜珈服",
            "运动内衣", "運動內衣", "文胸", "内衣", "內衣",
            "压缩裤", "壓縮裤", "紧身衣", "緊身衣", "legging", "泳衣", "泳镜", "泳帽",
            "蛋白粉", "氮泵", "肌酸", "維生素", "维生素", "保健品", "减肥药", "減肥藥",
        ],
        "heavy": ["沙袋", "负重背心", "負重背心", "阻力雪橇", "阻力雪撬",
                  "壶铃", "壺鈴", "杠铃片", "槓鈴片", "杠铃铃片", "槓鈴鈴片",
                  "哑铃片", "啞鈴片", "石磨", "实心球", "實心球"],
        "keep_files": ["儿童滑板车-m1688-2026-10-19-bikeparts-k1-表01.xlsx"],
    },
    # 戶外杯具 / 茶具 / 製茶製咖啡器具 (2026-10-20)
    # 價 20-150, 銷量 100+, 國際與國內大品牌都排除
    "cup": {
        "kwfile": "keywords_drinkware.txt",
        "prefix": "户外杯具茶具咖啡器具",
        "price": (20.0, 150.0),
        "min_sales": 100,
        "brand": [
            # ── 國際進口品牌 ──
            "Stanley", "STANLEY", "斯坦利",
            "Thermos", "THERMOS", "膳魔师", "膳魔師",
            "Zojirushi", "ZOJIRUSHI", "象印",
            "Yeti", "YETI", "Hydro Flask", "HYDROFLASK", "HYDROFLASK",
            "Owala", "OWALA", "Kinto", "KINTO",
            "Le Creuset", "S'well", "SWELL", "Corkcicle",
            "Takeya", "Nalgene", "NALGENE", "HydroJug", "RTIC",
            "Lifefactory", "MiiR", "Frank Green", "FRANK GREEN", "Bluel",
            "Hario", "HARIO", "V60", "Kalita", "KALITA",
            "Clever Dripper", "Chemex", "CHEMEX", "Comandante", "COMANDANTE",
            "Fellow", "Porlex", "PORLEX", "AeroPress", "AEROPRESS",
            "WACACO", "Nanopresso", "Staresso", "STARESSO",
            "Bialetti", "BIALETTI", "比乐蒂", "比樂蒂",
            "Timemore", "TIMEMORE", "泰摩", "AWADA", "BakerStone",
            "Mounthiker", "Mongdio", "Bincoo", "Nespresso", "NESPRESSO",
            "SMEG", "Dolce Gusto", "Tassimo", "Krups", "Bodum", "BODUM",
            "ZWILLING", "双立人", "雙立人", "WMF", "福腾宝", "福騰寶",
            "Baratza", "1Zpresso", "Tiamo", "Kanteen", "IKEA", "宜家",
            "Muji", "无印良品", "無印良品",
            "Snow Peak", "SNOWPEAK", "雪峰", "Naturehike", "NATUREHIKE", "挪客",
            "Coleman", "COLEMAN", "科勒曼", "科曼",
            "Tiger", "虎牌", "Lock&Lock", "LOCK&LOCK", "乐扣乐扣", "樂扣樂扣", "乐扣",
            "Oru Coffee", "Aarke",
            # ── 國內頭部品牌 ──
            "火枫", "火楓", "FIRE-MAPLE", "FIREMAPLE",
            "铠斯", "鎧斯", "KEITH", "拓村", "TOKON", "钛工匠", "鈦工匠",
            "千风茶道", "千風茶道", "THOUS WINDS", "山趣", "ShineTrip",
            "原始人", "探险者", "探險者", "克米特", "牧高笛", "Mobi Garden",
            "黑鹿", "BLACKDEER", "伯希和", "PELLIOT", "露得客", "户外家", "戶外家",
            "希诺", "希諾", "泰福高", "思宝", "思寶", "哈尔斯", "哈爾斯",
            "Haers", "FUGUANG", "富光", "益美", "炊大皇", "爱仕达", "愛仕達", "康巴赫",
            "小米", "Xiaomi", "XIAOMI", "米家", "MIJIA", "Mijia",
            "网易严选", "網易嚴選", "京东京造", "京東京造",
            "九阳", "九陽", "Joyoung", "苏泊尔", "蘇泊爾", "Supor", "美的", "Midea",
            # ── 茶具大牌 ──
            "金灶", "KAMJOVE", "恒福", "恆福", "容山堂", "尚明窑", "尚明風", "尚明风",
            "汉唐", "漢唐", "天福茗茶", "天福茶", "艺福堂", "藝福堂",
            "小罐茶", "大益", "八马", "八馬", "一厂一器", "一廠一器", "唐丰", "唐豐",
            # ── 咖啡品牌 ──
            "三顿半", "三頓半", "隅田川", "M Stand", "MStand", "Seesaw",
            "雀巢", "Nescafe", "NESCAFE", "Nescafé",
        ],
        "exclude": [
            # ── 茶葉 / 咖啡原料 / 食品 (本類是器具, 不是原料) ──
            "茶叶", "茶葉", "茶包", "茶饼", "茶餅", "普洱", "绿茶", "綠茶",
            "红茶", "紅茶", "白茶", "乌龙", "烏龍", "茉莉", "龙井", "龍井",
            "铁观音", "鐵觀音", "大红袍", "茶礼盒", "茶禮盒", "花草茶", "养生茶",
            "咖啡豆", "咖啡粉", "咖啡液", "挂耳", "掛耳", "速溶", "咖啡伴侣",
            "冷萃液", "浓缩咖啡液", "咖啡糖", "奶茶粉", "奶粉", "奶昔", "可可粉",
            # ── 電器 / 電動 (跨境運輸 + 台灣電壓法規) ──
            "电动", "電動", "电动磨豆", "电动咖啡", "咖啡机", "咖啡機", "意式机", "意式機",
            "电磨", "电热水壶", "電熱水壺", "烧水壶", "燒水壺", "恒温", "恆溫",
            "电磁炉", "電磁爐", "电陶炉", "破壁机", "榨汁机", "料理机", "烤箱",
            "微波炉", "电饭煲", "养生壶", "电煮", "电炖", "电热", "電熱",
            "加热", "加熱", "充电", "充電", "插电", "插電", "数显", "顯示",
            "智能", "APP控制", "电子烟",
            # ── 燃氣 / 爐具 (戶外煮茶常見, 但本身是禁賣燃氣品) ──
            "卡式炉", "卡式爐", "瓦斯罐", "气罐", "氣罐", "酒精炉", "露营炉", "露營爐",
            "柴火炉", "烧烤炉", "燒烤爐", "烤盘", "烤盤", "烤网", "烤網", "火盆",
            "取暖器", "暖风机", "炉具", "野餐炉",
            # ── 刀具 (2026-10-20 使用者確認: 刀具可以, 不排除) ──
            # ── 煙酒 ──
            "香烟", "煙香", "烟嘴", "煙嘴", "烟灰缸", "煙灰缸", "烟盒", "煙盒",
            "打火机", "打火機", "酒壶", "酒壺", "醒酒器", "分酒器", "酒嘴", "酒塞",
            "酒漏", "威士忌", "洋酒", "红酒杯", "紅酒杯", "鸡尾酒", "雞尾酒",
            "调酒", "調酒", "bartender", "扎啤", "啤酒杯", "酒吧", "婚庆", "婚慶",
            # ── 藥膳 / 保健 ──
            "保健", "养生", "養生", "中药", "中藥", "药膳", "藥膳", "足浴", "泡脚",
            "减肥", "減肥", "助眠", "精油",
            # ── 兒童 / 嬰幼兒 (台灣要求安全標章, 風險高) ──
            "儿童", "兒童", "婴儿", "嬰兒", "宝宝", "寶寶", "幼儿", "幼兒", "卡通",
            # ── 廚房非杯具 / 清潔 / 收納 ──
            "洗洁精", "清洁剂", "清潔劑", "去污", "冰箱", "保鲜", "保鮮",
            "饭碗", "碗碟", "餐盘", "圍裙", "围裙", "抹布",
            # ── 裝飾 / 擺件 / 風水 ──
            "摆件", "擺件", "装饰", "裝飾", "风水", "風水", "招财", "招財",
            "貔貅", "挂件", "掛件", "茶宠", "茶寵",
            # ── 二手 / 代工 / 清倉 ──
            "二手", "中古", "回收", "闲置", "閒置", "贴牌", "貼牌", "代工",
            "定制logo", "定製logo", "贴标", "貼標", "库存", "庫存", "清仓", "清倉",
            # ── 服飾 ──
            "衣服", "服裝", "服饰", "服飾", "T恤", "裤子", "褲子", "袜子", "襪子",
            # ── 其他非本類 ──
            "宠物", "寵物", "狗", "猫", "貓",
            "自行车", "自行車", "电动车", "電動車", "汽车用品", "車用",
            "医疗器械", "醫療器械", "输液", "點滴", "纪念品",
        ],
        # 20-150 價帶沒有重量風險 (保溫壺/玻璃壺除外), 不設超重詞
        "heavy": [],
        "keep_files": [],
    },
    # 保溫飯盒 / 保溫湯罐 / 保溫箱 (2026-10-20)
    # 戶外用途: 保生鮮 / 保冰塊 / 帶飯 / 燉湯
    # 價 20-220, 銷量 200+, 國際與國內大品牌都排除
    "insul": {
        "kwfile": "keywords_insulated_box.txt",
        "prefix": "保温饭盒汤罐保温箱",
        "price": (20.0, 220.0),
        "min_sales": 200,
        # 只排除「已交付最終總表」— 舊任務被篩掉的保溫類品項要能重新採集
        "dedup_mode": "delivered",
        "brand": [
            # ── 國際進口品牌 (保溫箱/冰桶/飯盒) ──
            "Stanley", "STANLEY", "斯坦利", "史丹利",
            "Thermos", "THERMOS", "膳魔师", "膳魔師",
            "Zojirushi", "ZOJIRUSHI", "象印", "ZEBRA", " zebra",
            "Tiger", "TIGER", "虎牌",
            "Lock&Lock", "LOCK&LOCK", "乐扣乐扣", "樂扣樂扣", "乐扣",
            "Coleman", "COLEMAN", "科勒曼", "科曼",
            "Igloo", "IGLOO", "艾格鲁",
            "RTIC", "YETI", "Hydro Flask", "HYDROFLASK",
            "Playmate", "Eliminator", "Rubbermaid", "Rubbermaid",
            "Pelican", "PELICAN", "海霸", "Dometic", "ALPS Mountaineering",
            "Euhsy", "ICEMULE", "Grizzly", "Ozark", "OZARK",
            "Icepack", "ICEPACK", "爱思比", "愛思比", "爱路普", "愛路普",
            "Sawafuji", "SAHARA", "SAKURA",
            "Decathlon", "DECATHLON", "迪卡侬",
            # ── 國際戶外大牌 ──
            "Naturehike", "NATUREHIKE", "挪客",
            "Snow Peak", "SNOWPEAK", "雪峰",
            "Mobi Garden", "MOBIGARDEN", "牧高笛",
            "KingCamp", "KINGCAMP", "Blackdeer", "BLACKDEER", "黑鹿",
            "Helinox", "HELINOX", "黑诺", "黑諾",
            "Sea to Summit", "SeaToSummit", "Therm-a-Rest", "ThermARest",
            "Big Agnes", "Marmot", "MARMOT", "山之客",
            "The North Face", "TNF", "Jack Wolfskin", "狼爪",
            "CAMEL", "骆驼", "駱駝", "CRAGHOPPERS",
            "Klymit", "KLYMIT", "Intex", "INTEX", "3F UL Gear",
            "CAMPFIRE", "Campfire", "闲风蓝", "閒風藍",
            "露营先生", "露營先生", "TILLAK", "SWS WELLHIKE", "WELLHIKE",
            "DOD", "露得客", "户外家", "戶外家", "户外家户外",
            # ── 國內頭部戶外品牌 ──
            "原始人", "探险者", "探險者", "克米特",
            "伯希和", "PELLIOT", "火枫", "火楓", "FIRE-MAPLE", "FIREMAPLE",
            "铠斯", "鎧斯", "KEITH", "拓村", "TOKON", "钛工匠", "鈦工匠",
            "探路者", "TOREAD", "凯乐石", "KAILAS", "挪客户外",
            # ── 國內頭部保溫器皿品牌 ──
            "希诺", "希諾", "泰福高", "思宝", "思寶", "哈尔斯", "哈爾斯",
            "Haers", "FUGUANG", "富光", "益美", "道奇", "稻草人",
            "炊大皇", "爱仕达", "愛仕達", "康巴赫", "双喜", "雙喜",
            "TYESO", "泰硕", "泰碩", "尚明", "容山堂", "金灶",
            "茶花", "CHAHUA",
            # ── 頭部電商 / 家電品牌 ──
            "小米", "Xiaomi", "XIAOMI", "米家", "MIJIA", "Mijia",
            "网易严选", "網易嚴選", "京东京造", "京東京造",
            "九阳", "九陽", "Joyoung", "苏泊尔", "蘇泊爾", "Supor",
            "美的", "Midea", "小熊", "Bear", "北鼎", "摩飞", "摩飛",
        ],
        "exclude": [
            # ── 電動 / 插電 / 蓄電 (跨境運輸 + 台灣電壓法規) ──
            "电动", "電動", "电热", "電熱", "加热", "加熱", "充电", "充電",
            "插电", "插電", "恒温", "恆溫", "数显", "顯示", "智能",
            "电煮", "電煮", "电炖", "電燉", "usb", "USB", "电池", "電池",
            "led", "LED", "显示屏", "顯示屏", "液晶", "温控", "溫控",
            # ── 醫療 / 保健 (冰袋、急救箱…) ──
            "医用", "醫用", "医疗", "醫療", "急救", "保健", "穴位",
            "婴儿", "嬰兒", "宝宝", "寶寶", "幼儿", "幼兒", "儿童", "兒童",
            "血糖", "胰岛素", "霧化", "雾化", "呼吸机", "呼吸機",
            # ── 食品 / 食材 (本類是容器, 不是食物) ──
            "预制菜", "預製菜", "熟食", "卤味", "滷味", "速冻", "速凍",
            "冷冻食品", "冷凍食品", "便当包", "便當包", "饭盒包", "飯盒包",
            "烘焙", "面包", "麵包", "点心", "點心", "寿司", "壽司",
            "午餐肉", "火腿", "零食", "冻品", "肉卷", "香肠", "臘腸",
            # 注意: 「水果 / 生鲜 / 海鲜」不能排除 — 「水果保鲜箱」「海鲜保鲜箱」是本類正品
            # 「蛋糕 / 奶茶店 / 咖啡店 / 微波炉」也不排除 —
            #   奶茶保温桶是保溫桶, 「可微波炉」是保溫飯盒的功能描述
            # ── 一次性 / 外賣 / 商用配送 ──
            "一次性", "打包盒", "外卖", "外賣", "快餐盒", "餐盒一次性",
            "配送箱", "送餐箱", "外卖箱", "冷藏车", "冷藏車", "保温车",
            "保温车", "美团", "众包", "騎手", "骑手",
            # ── 廚房炊具 / 電器 ──
            "电饭煲", "電飯煲", "电压力锅", "電壓力鍋",
            "烤箱", "电磁炉", "電磁爐", "破壁机", "榨汁", "水龙头", "水龍頭",
            "餐桌", "椅子", "餐椅", "折叠桌", "折叠椅", "野餐桌", "置物架",
            "展示架", "收纳架", "收納架", "货架", "貨架",
            # ── 燃氣 / 煙酒 ──
            "卡式炉", "卡式爐", "瓦斯", "气罐", "氣罐", "燃气", "燃氣",
            "打火机", "打火機", "烟嘴", "煙嘴", "酒壶", "酒壺", "醒酒器",
            "红茶", "紅茶", "绿茶", "綠茶", "咖啡豆", "咖啡粉",
            # ── 服飾 / 其他非本類 ──
            "衣服", "服裝", "服饰", "服飾", "T恤", "裤子", "褲子", "袜子", "襪子",
            "运动鞋", "運動鞋", "宠物", "寵物", "狗", "猫", "貓",
            "自行车", "自行車", "电动车", "電動車", "汽车用品", "車用",
            "二手", "中古", "回收", "闲置", "閒置", "贴牌", "貼牌", "代工",
            "定制logo", "定製logo", "贴标", "貼標", "库存", "庫存", "清仓", "清倉",
            "装饰", "裝飾", "摆件", "擺件", "纪念品", "紀念品",
        ],
        "heavy": [],
        "keep_files": [],
    },
}

# 機車後尾箱 / 機車後靠背 (2026-10-21) — 設定集中在 moto_rules.py 維護
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
from moto_rules import COLLECT_CFG as _MOTO_COLLECT_CFG
CONFIG.update(_MOTO_COLLECT_CFG)

# 舞蹈鞋 / 攀岩鞋 / 溯溪鞋 (2026-10-28) — 設定集中在 shoe_rules.py 維護
from shoe_rules import COLLECT_CFG as _SHOE_COLLECT_CFG
CONFIG.update(_SHOE_COLLECT_CFG)

# 居家垃圾桶 (2026-10-28) — 設定集中在 trash_rules.py 維護
from trash_rules import COLLECT_CFG as _TRASH_COLLECT_CFG
CONFIG.update(_TRASH_COLLECT_CFG)

# 居家+店铺装潢 垃圾桶 第二彈 (2026-10-28) — trash2_rules.py
from trash2_rules import COLLECT_CFG as _TRASH2_COLLECT_CFG
CONFIG.update(_TRASH2_COLLECT_CFG)

# 居家墙面置物架 / 挂架 / 挂杆 (2026-11-02) — wall_rules.py
from wall_rules import COLLECT_CFG as _WALL_COLLECT_CFG
CONFIG.update(_WALL_COLLECT_CFG)

# 宠物窝 / 宠物床 / 窝垫 (2026-11-04) — pet_bed_rules.py
from pet_bed_rules import COLLECT_CFG as _PETBED_COLLECT_CFG
CONFIG.update(_PETBED_COLLECT_CFG)

# 宠物喂食 (宠物碗/喂食器/喂水器) (2026-11-06) — pet_feed_rules.py
from pet_feed_rules import COLLECT_CFG as _PETFEED_COLLECT_CFG
CONFIG.update(_PETFEED_COLLECT_CFG)

# 运动头盔 (骑行/滑板/轮滑/滑雪/登山攀岩/球类) (2026-11-17) — helmet_rules.py
from helmet_rules import COLLECT_CFG as _HELMET_COLLECT_CFG
CONFIG.update(_HELMET_COLLECT_CFG)

# 頁面狀態分類 + 風控處置決策表 (2026-11-09, 移植自 GitHub superjack2050/1688-cli)
import v3_page_state as PS

# 居家桌板 / 便携桌板 / 床桌 (2026-11-25) — desktb_rules.py
# 價 25-250; 銷量 200-2000 (區間); 不要國際大品牌
from desktb_rules import COLLECT_CFG as _DESKTB_COLLECT_CFG
CONFIG.update(_DESKTB_COLLECT_CFG)

CFG = CONFIG[_KWSET]
# 品類可自訂日期 (新開品類用當天日期); 沒有此鍵的品類沿用預設 DATE
DATE = CFG.get("date", DATE)
KW_FILE = SCRIPT_DIR.parent / CFG["kwfile"]
PRICE_MIN, PRICE_MAX = CFG["price"]
MIN_SALES = CFG["min_sales"]
# 銷量上限 (可選): 沒有此設定的品類維持原本的「只有下限」行為
MAX_SALES = CFG.get("max_sales")
BRAND_EXCLUDE = CFG["brand"]
# 平台禁賣品 (護目鏡/防護眼鏡/滑雪鏡/游泳鏡…) — 所有品類一律擋 (共用引擎層定義)
from v3_rules import BANNED as _BANNED_EXCLUDE
EXCLUDE_KEYWORDS = list(CFG["exclude"]) + _BANNED_EXCLUDE
# regex 版排除 (2026-11-06 新增): 支援「後接特定字就不算」
# 例: 猫粮(?!碗|盆|盒|包|机|器) → 「猫粮碗」是碗不是糧, 不可殺
EXCLUDE_REGEX = [re.compile(p) for p in (CFG.get("exclude_regex") or [])]
HEAVY_ITEM_KW = CFG["heavy"]
TABLE_PREFIX = CFG["prefix"]
TARGET_PER_TABLE = 100

# ── 節奏 (v3: 沒有 document 載入, 請求量比 v2 少一個數量級) ──────────
REQ_GAP_MIN, REQ_GAP_MAX = _envf("M1688_REQ_GAP_MIN", 4.0), _envf("M1688_REQ_GAP_MAX", 8.0)          # 同一詞內翻頁的間隔
KW_GAP_MIN, KW_GAP_MAX = _envf("M1688_KW_GAP_MIN", 12.0), _envf("M1688_KW_GAP_MAX", 25.0)         # 詞與詞的間隔
COOL_DOWN_EVERY_MIN, COOL_DOWN_EVERY_MAX = _envi("M1688_CD_EVERY_MIN", 8), _envi("M1688_CD_EVERY_MAX", 12)
COOL_DOWN_SEC_MIN, COOL_DOWN_SEC_MAX = _envf("M1688_CD_SEC_MIN", 60), _envf("M1688_CD_SEC_MAX", 150)
LONG_PAUSE_CHANCE = _envf("M1688_LONG_PAUSE_CHANCE", 0.15)
LONG_PAUSE_SEC_MIN, LONG_PAUSE_SEC_MAX = _envf("M1688_LONG_SEC_MIN", 45), _envf("M1688_LONG_SEC_MAX", 90)
AUTO_SAVE_EVERY = 2
HARD_TIMEOUT_SEC = 180 * 60
RETRY_PER_PAGE = 3                  # 單頁失敗重試次數 (伺服器偶發逾時)
RETRY_BACKOFF_MIN, RETRY_BACKOFF_MAX = 4.0, 9.0

DONE_FILE = Path(tempfile.gettempdir()) / f"1688_done_v3_{_KWSET}.txt"

LOG_LINES = []

# ── 看門狗 ────────────────────────────────────────────────────────────
# page.evaluate() 沒有逾時機制, 遇到渲染程序被卡住 (分頁卡頓/正在導頁) 會永久阻塞,
# JS 內部的 AbortController 也救不了 (JS 根本沒在跑)。2026-10-20 實測卡死 >10 分鐘。
# 解法: 監看「最後一次拿到 API 回應」的時間, 超過門檻就主動結束程式。
# 批次檔每 AUTO_SAVE_EVERY 個詞自動存檔, 所以最多只損失那一段。
LAST_HIT_TS = time.time()
STALL_TIMEOUT_SEC = 300
_wd_lock = threading.Lock()


def touch_hit():
    global LAST_HIT_TS
    with _wd_lock:
        LAST_HIT_TS = time.time()


def start_watchdog():
    def _run():
        while True:
            time.sleep(15)
            with _wd_lock:
                idle = time.time() - LAST_HIT_TS
            if idle > STALL_TIMEOUT_SEC:
                try:
                    log(f"[WATCHDOG] 連續 {int(idle)}s 沒有任何 API 回應, 判定卡死, 結束程式"
                        f" (已收資料都已存檔, 可續跑)")
                except Exception:
                    pass
                os._exit(3)
    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return t


def log(msg):
    line = f"  {msg}"
    print(line, flush=True)
    LOG_LINES.append(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def jitter(base_sec):
    """拆段 sleep, 總時長精確等於 base (2026-10-18 修正過的歸一化版本)

    ★ 2026-11-17 修正: 拆段 sleep 期間**定期 touch_hit()**。
      原本只在 sleep 前後各 touch 一次 → 若單次等待超過 STALL_TIMEOUT_SEC(300s)
      (詞間隔 + 長休 + 冷卻 疊加可達 750s), 看門狗會在睡眠期間誤判「卡死」而殺掉程式。
      2026-11-17 實測因此連續兩輪被誤殺 (exit=3)。
      → 每 60 秒補一次心跳, 讓等待期間不被當成卡死。
    """
    if base_sec <= 0:
        return 0.0
    n = random.randint(3, 6)
    w = [random.uniform(0.55, 1.45) for _ in range(n)]
    s = sum(w)
    for x in w:
        seg = max(1.0, base_sec * x / s)
        # 拆成 ≤60s 的小段, 每段結束補一次心跳
        left = seg
        while left > 0:
            step = min(60.0, left)
            time.sleep(step)
            left -= step
            touch_hit()


def load_done_keywords():
    try:
        return set(l.strip() for l in DONE_FILE.read_text(encoding="utf-8").split("\n") if l.strip())
    except Exception:
        return set()


def mark_done(kw):
    try:
        with open(DONE_FILE, "a", encoding="utf-8") as f:
            f.write(kw + "\n")
    except Exception:
        pass
    # 2026-11-04: 另存「按日期」記錄, 讓每日詞數預算能算到非本守門模組採集的日子
    try:
        bydate = TEMP_DIR / f"1688_done_v3_{_KWSET}_bydate.txt"
        with open(bydate, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d')}\t{kw}\n")
    except Exception:
        pass


def load_self_batch_offerids():
    """本品類既有批次檔裡的 offerId — 續跑時不要重複採集同一批商品

    背景: 採集程序被中止後要續跑, 如果只靠跨任務去重池 (delivered 模式不看池),
    續跑會把前一輪已收的品項再收一次 (浪費請求 + 產生重複列)。
    """
    from openpyxl import load_workbook
    oids = set()
    prefix = CFG["prefix"]
    for f in EXCEL_DIR.glob(f"{prefix}-m1688-{DATE}-bikeparts-*.xlsx"):
        if "DISCARD" in f.name or "_test" in f.name:
            continue
        try:
            wb = load_workbook(f, read_only=True)
            ws = wb.worksheets[0]
            hdr = [str(c.value) if c.value else "" for c in next(ws.iter_rows(max_row=1))]
            if "商品链接" not in hdr:
                wb.close()
                continue
            ci = hdr.index("商品链接")
            for r in ws.iter_rows(min_row=2, values_only=True):
                u = str(r[ci] or "")
                if "/offer/" in u:
                    oids.add(u.split("/offer/")[1].split(".html")[0])
            wb.close()
        except Exception as e:
            log(f"[WARN] 讀批次檔失敗 {f.name}: {e}")
    if oids:
        log(f"[OK] 本類別既有批次 offerId: {len(oids)} 个 (續跑時不重複採)")
    return oids


def load_delivered_offerids():
    """已交付最終總表裡的 offerId — 這些使用者已經上架/看過, 不該重複採

    為什麼需要這個模式:
    跨任務去重池 (_existing_offerids_2026-08-27.txt) 會把「曾經採集過」的品項全部擋掉,
    包括那些在舊任務複審階段被篩掉的 (品牌不符 / 非本類)。
    對新品類來說那些品項正是要的, 所以改成只擋「已交付總表」裡的。
    """
    from openpyxl import load_workbook
    oids = set()
    files = list(EXCEL_DIR.glob("*最終總表*.xlsx")) + list(EXCEL_DIR.glob("*今日總表*.xlsx"))
    for f in files:
        try:
            wb = load_workbook(f, read_only=True)
            ws = wb.worksheets[0]
            hdr = [str(c.value) if c.value else "" for c in next(ws.iter_rows(max_row=1))]
            if "商品链接" not in hdr:
                wb.close()
                continue
            ci = hdr.index("商品链接")
            for r in ws.iter_rows(min_row=2, values_only=True):
                u = str(r[ci] or "")
                if "/offer/" in u:
                    oids.add(u.split("/offer/")[1].split(".html")[0])
            wb.close()
        except Exception as e:
            log(f"[WARN] 讀已交付總表失敗 {f.name}: {e}")
    log(f"[OK] 已交付總表 offerId: {len(oids)} 个 ({len(files)} 个檔)")
    return oids


def load_existing_offerids():
    mode = CFG.get("dedup_mode", "global")
    if mode == "delivered":
        log("[DEDUP] 模式 = 只排除「已交付最終總表」(舊任務被篩掉的品項可重新採集)")
        oids = load_delivered_offerids()
        oids |= load_self_batch_offerids()
        return oids
    try:
        with open(EXISTING_OIDS_FILE, encoding="utf-8") as f:
            oids = set()
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    oids.add(line)
        log(f"[OK] 跨任务去重池: {len(oids)} 个 offerId")
        return oids
    except Exception:
        log(f"[WARN] 跨任务去重池不存在: {EXISTING_OIDS_FILE}")
        return set()


def save_oid_pool(new_oids):
    """寫入跨任務去重池 — 已存在的 id 不重複追加"""
    try:
        have = set()
        if EXISTING_OIDS_FILE.exists():
            with open(EXISTING_OIDS_FILE, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        have.add(line)
        todo = [o for o in sorted(new_oids) if o not in have]
        if not todo:
            return
        with open(EXISTING_OIDS_FILE, "a", encoding="utf-8") as f:
            for o in todo:
                f.write(o + "\n")
    except Exception as e:
        log(f"[WARN] 寫去重池失敗: {e}")


def load_keep_offerids():
    """同品類舊 DOM 批次檔裡的 offerId — 這些要保留, 不能被跨任務去重池擋掉"""
    keep = set()
    from openpyxl import load_workbook
    for fn in CFG["keep_files"]:
        p = EXCEL_DIR / fn
        if not p.exists():
            continue
        try:
            wb = load_workbook(p, read_only=True)
            ws = wb.worksheets[0]
            for row in ws.iter_rows(min_row=2, values_only=True):
                u = row[9] if len(row) > 9 else ""
                if u and "/offer/" in str(u):
                    keep.add(str(u).split("/offer/")[1].split(".html")[0])
            wb.close()
        except Exception as e:
            log(f"[WARN] 讀 keep 檔失敗 {fn}: {e}")
    if keep:
        log(f"[OK] 同類舊檔白名單: {len(keep)} 筆 (會保留)")
    return keep


# ── mtop 簽名 + 請求 ────────────────────────────────────────────────
# ★★ 2026-11-06 重大修正: 請求格式必須與「真人網頁」完全一致 ★★
# 實測 (capture_live_request.py 抓真人搜尋流量):
#   真人: {"sortType":"pop","keywords":KW,"filtId":"","p4pCount":"4","appName":"wap",
#          "appKey":DATA_APP_KEY,"beginPage":1,"pageSize":20}
#   舊版採集器 (錯): pageSize=60 / sortType=price_asc / 多帶 priceStart+priceEnd
#   → 瀏覽器永遠不會送 pageSize=60 與 priceStart/priceEnd
#   → 1688 前端改版後, 舊指紋一眼就是「機器」→ 風控 (2026-11-06 連撞兩次驗證)
#
# 真人一次搜尋會發 9 個請求 (4 種不同接口), 採集器只發 1 種 → 請求圖是孤島
#   widget.get / hamlet.tql / hotkeyword.getHotKeyWords / suggestion.getSuggestWords
#   / sn.common.getStatisticsData / seo.offer.get(搜尋) / widget.get / hotkeyword / hamlet
# → 加入「擬人流量」: 每個詞依真人順序補上輔助請求

PAGE_SIZE_HUMAN = 20          # 真人頁面固定 20
SORT_TYPE_HUMAN = "pop"       # 頁面預設「综合」(熱度) — 2026-11-06 實測
DATA_APP_KEY = "1772295d201e06bd6d56a6ed5cfd252f"


def _mtop_url(api, data_obj, token, callback="mtopjsonp1", jsv="2.5.1"):
    data = json.dumps(data_obj, ensure_ascii=False, separators=(",", ":"))
    t = str(int(time.time() * 1000))
    sign = hashlib.md5(f"{token}&{t}&{MTOP_APP_KEY}&{data}".encode("utf-8")).hexdigest()
    return ("https://h5api.m.1688.com/h5/" + api + "/1.0/"
            f"?jsv={jsv}&appKey={MTOP_APP_KEY}&t={t}&sign={sign}"
            f"&api={api}&v=1.0&type=jsonp&dataType=jsonp"
            f"&jsonpIncPrefix=mtopjsonp&callback={callback}&data={quote(data, safe='')}")


def build_mtop_url(token, kw, begin_page, page_size=None, callback="mtopjsonp1"):
    """搜尋請求 — payload 與「真人網頁」完全一致

    ★ 2026-11-06 實測 (capture_live_request.py / capture_price_filter.py):
      頁面點「价格」排序時送:
        {"sortType":"price","filtId":"","keywords":KW,"descendOrder":"false",
         "appName":"wap","appKey":DATA_APP_KEY,"beginPage":N,"pageSize":20}
      頁面預設「综合」排序時送:
        {"sortType":"pop",...,"p4pCount":"4",...,"pageSize":20}
      → 舊版採集器送的是 sortType=price_asc + p4pCount + pageSize=60,
        瀏覽器永遠不會送這種混合體 → 一眼就是機器 → 風控 (2026-11-06 連撞)
    """
    fast = os.environ.get("M1688_FAST", "").strip() == "1"
    if fast:
        # 舊的激進模式 (只供除錯/量產對照, 預設不用)
        page_size = page_size or PAGE_SIZE
        data = {
            "sortType": SORT_TYPE,
            "keywords": kw,
            "filtId": "",
            "p4pCount": "4",
            "appName": "wap",
            "appKey": DATA_APP_KEY,
            "beginPage": begin_page,
            "pageSize": page_size,
            "priceStart": f"{int(PRICE_MIN)}",
            "priceEnd": f"{int(PRICE_MAX)}",
        }
    else:
        # ★★★ 已驗證配置 (2026-10-20 ~ 10-30: 25 批次 / 4,882 筆 / **0 次硬風控**) ★★★
        #   sortType=price_asc → 後端會落到「預設排序(銷量傾向)」→ 命中率 14~41/60
        #   (pop 只有 4/60 — 用 pop 等於把產出砍到 1/4, 2026-11-06 誤改過一次)
        #   pageSize=60 → 實測 60/100 都可 (200 會被砍成 60)
        #   priceStart/priceEnd → 伺服器直接壓進價帶 (省本地過濾, 命中率大增)
        # ⚠ 這些參數「瀏覽器不會送」, 但歷史證據: 25 批次 0 硬風控 → payload 從來不是風控成因,
        #   成因是「節奏」(一次跑太多詞 / 並行 / 觀察期硬幹)
        page_size = page_size or PAGE_SIZE
        data = {
            "sortType": SORT_TYPE,
            "keywords": kw,
            "filtId": "",
            "p4pCount": "4",
            "appName": "wap",
            "appKey": DATA_APP_KEY,
            "beginPage": begin_page,
            "pageSize": page_size,
            "priceStart": f"{int(PRICE_MIN)}",
            "priceEnd": f"{int(PRICE_MAX)}",
        }
    return _mtop_url(MTOP_API, data, token, callback)


# 擬人流量: 真人搜尋時會跟著發的輔助請求 (payload 全部來自實測)
def build_human_traffic_urls(token, kw):
    """回傳 [(說明, url), ...] — 依真人頁面的實際順序

    ★ 節奏考量 (2026-11-06): 真人一次搜尋發 9 個請求, 但採集器若照抄
      會讓「每詞 8 個請求」, 總量是舊版 (3/詞) 的 2.7 倍 → 帳號更快被盯上。
      → 預設只留「前 1 個 (suggest) + 後 1 個 (hotkey)」= 每詞 +2 個請求,
        既保留「多種接口交錯」的特徵, 又把總量壓回去。
      → M1688_HUMAN_TRAFFIC=full → 照抄真人全部 5+2 個
      → M1688_HUMAN_TRAFFIC=0    → 完全關閉 (回到純搜尋請求)
    """
    mode = os.environ.get("M1688_HUMAN_TRAFFIC", "0").strip()
    out = []
    if mode == "full":
        out.append(("suggest", _mtop_url(
            "mtop.1688.suggestion.common.getSuggestWords",
            {"type": "offer", "appName": "wap", "appKey": DATA_APP_KEY, "keywords": kw},
            token)))
        out.append(("hotkey", _mtop_url(
            "mtop.1688.hotkeyword.common.getHotKeyWords",
            {"keywords": kw, "rescnt": 60, "appName": "wap",
             "appKey": DATA_APP_KEY, "source": "MAIN"},
            token)))
        out.append(("statistics", _mtop_url(
            "mtop.1688.sn.common.getStatisticsData",
            {"keywords": kw, "appName": "wap", "appKey": DATA_APP_KEY},
            token)))
    elif mode == "light":
        out.append(("suggest", _mtop_url(
            "mtop.1688.suggestion.common.getSuggestWords",
            {"type": "offer", "appName": "wap", "appKey": DATA_APP_KEY, "keywords": kw},
            token)))
    return out


def build_human_traffic_after(token, kw):
    """搜尋之後真人還會補的請求"""
    mode = os.environ.get("M1688_HUMAN_TRAFFIC", "0").strip()
    out = []
    if mode == "full":
        out.append(("widget", _mtop_url(
            "mtop.1688.wap.widget.get",
            {"keywords": kw, "brandPid": "112017_0003", "appName": "brandWapRender"},
            token)))
        out.append(("hotkey2", _mtop_url(
            "mtop.1688.hotkeyword.common.getHotKeyWords",
            {"keywords": kw, "rescnt": 60, "appName": "wap",
             "appKey": DATA_APP_KEY, "source": "MAIN"},
            token)))
    elif mode == "light":
        out.append(("hotkey2", _mtop_url(
            "mtop.1688.hotkeyword.common.getHotKeyWords",
            {"keywords": kw, "rescnt": 60, "appName": "wap",
             "appKey": DATA_APP_KEY, "source": "MAIN"},
            token)))
    return out


def fetch_mtop(page, url):
    """在頁面 context 內 fetch — 沿用 Chromium TLS 指紋 + cookie
    帶 AbortController 超時 (35s), 否則網路卡住會永久阻塞整支腳本"""
    touch_hit()
    try:
        return page.evaluate("""async (u) => {
            const ctrl = new AbortController();
            const timer = setTimeout(() => ctrl.abort(), 35000);
            try {
                const r = await fetch(u, {credentials: 'include', signal: ctrl.signal});
                const t = await r.text();
                clearTimeout(timer);
                return {ok: true, status: r.status, text: t};
            } catch (e) {
                clearTimeout(timer);
                return {ok: false, status: -1, text: String(e)};
            }
        }""", url)
    except Exception as e:
        # 2026-11-04 新增: 風控會在「請求進行中」把分頁導到驗證頁 → execution context 被銷毀
        #   → page.evaluate 不回應也不拋錯, 會懸空到看門狗 (300s) 才殺程式
        #   舊版只在「回應文字」裡找 访问被拒绝/_____tmd_____ , 抓不到這種形式
        # 處置: 檢查分頁現在的 URL, 若已被導到風控/驗證/登入頁 → 回報 HARD_BLOCK 讓上層立刻中止
        try:
            now_url = page.url or ""
        except Exception:
            now_url = ""
        if any(t in now_url for t in ("_____tmd_____", "punish", "sec_verify",
                                      "login.taobao.com", "login.1688.com")):
            return {"ok": False, "status": -2,
                    "text": f"HARD_BLOCK:頁面已被導至風控/驗證頁 {now_url[:140]}"}
        # 分頁被使用者關掉/導走 -> 交給 ensure_page 重建, 這裡只回報失敗
        return {"ok": False, "status": -1, "text": f"EVAL_ERR: {e}"}


# 專用採集分頁的識別標記 (避免動到使用者自己正在用的分頁)
COLLECT_URL = "https://m.1688.com/?_mc=1"
# own_tab (2026-11-04): 只用自己開的分頁, 不重用使用者的分頁
#   實測: 使用者操作同一個分頁 → context 被銷毀 → 採集懸空 300 秒
#   舊品類預設 False (維持原行為), 新品類可設 True
OWN_TAB = CFG.get("own_tab", False)


def wait_page_settled(page, tries=20, interval=0.5):
    """等分頁 URL 穩定後再發請求 (2026-11-04)

    1688 手機版 SPA 載入後會改網址 (加 callByHgJs / __removesafearea__ / src_cna…),
    在導頁期間發出的 fetch 會讓 execution context 被銷毀 → 請求懸空或報錯
    (實測: 新分頁的第一次請求 499 秒 → 加 focus emulation 後 56 秒)
    解法: 輪詢 page.url, 連續幾次不變才算穩定
    """
    last, stable = None, 0
    for _ in range(tries):
        try:
            cur = page.url or ""
        except Exception:
            cur = ""
        if cur and cur == last:
            stable += 1
            if stable >= 3:
                return cur
        else:
            stable = 0
            last = cur
        time.sleep(interval)
    return last


def enable_focus_emulation(ctx, page):
    """讓採集分頁「以為自己有焦點」— 避免 Chrome 背景分頁凍結 (2026-11-04)

    實測徵兆: page.evaluate 既不回應也不拋錯, 懸空 300 秒 → 看門狗殺人
    原因: CDP 開的新分頁不會成為瀏覽器 active tab → document.visibilityState=hidden
          → 5 分鐘後 Chrome 凍結該分頁, JS 完全停擺
          → fetch 永不完成、AbortController 的 setTimeout 也不觸發
    解法: Emulation.setFocusEmulationEnabled → 頁面視為 focused, 不凍結、不節流
          (不必重啟 Chrome, 也不搶使用者的焦點)
    """
    try:
        s = ctx.new_cdp_session(page)
        s.send("Emulation.setFocusEmulationEnabled", {"enabled": True})
        log("  [FOCUS] 已啟用 focus emulation (防止背景分頁凍結)")
        return True
    except Exception as e:
        log(f"  [FOCUS] focus emulation 設定失敗: {str(e)[:90]}")
        return False


def _capture_foreground():
    """記住目前 OS 前景視窗的 handle (開分頁會把 Chrome 拉到前景)"""
    try:
        import ctypes
        return ctypes.windll.user32.GetForegroundWindow()
    except Exception:
        return None


def _restore_foreground(hwnd):
    """把焦點還給使用者原本在用的視窗

    2026-11-04 實測: ctx.new_page() 會把 Chrome 視窗拉到前景,
    使用者會看到「我開的 1688 實例自己打開了」。

    注意: Windows 的「前景鎖」會擋下跨進程 SetForegroundWindow
    (尤其目標視窗是工作管理員等高權限程式時會直接失敗),
    → 用 AttachThreadInput 把執行緒輸入佇列接上再設前景 (標準繞法)。
    """
    if not hwnd:
        return
    try:
        import ctypes
        u = ctypes.windll.user32
        k = ctypes.windll.kernel32
        target_thread = u.GetWindowThreadProcessId(hwnd, None)
        current_thread = k.GetCurrentThreadId()
        attached = False
        if target_thread and target_thread != current_thread:
            attached = bool(u.AttachThreadInput(current_thread, target_thread, True))
        try:
            u.ShowWindow(hwnd, 9)          # SW_RESTORE
            u.BringWindowToTop(hwnd)
            u.SetForegroundWindow(hwnd)
            u.SetActiveWindow(hwnd)
        finally:
            if attached:
                u.AttachThreadInput(current_thread, target_thread, False)
        # 仍失敗時的最後手段 (跨進程強制切換)
        if u.GetForegroundWindow() != hwnd:
            u.SwitchToThisWindow(hwnd, True)
    except Exception:
        pass


def new_background_page(ctx, url):
    """開一個採集用分頁, 並把焦點還給使用者 (不搶走他正在用的視窗)

    2026-11-04 實測:
      · ctx.new_page() 會讓 Chrome 選中新分頁 → 視窗自動跳到前景,
        使用者會看到「我開的 1688 實例自己打開了」。
      · CDP Target.createTarget(background=True) 雖然真的能背景建立 target,
        但 **Playwright 不會接手** (ctx.pages 不會出現那個分頁) → 不可用。
      → 唯一可行解: ctx.new_page() + 事後把 OS 焦點還回去。
      背景/非選中分頁有被 Chrome 凍結的風險 → 必須搭配 focus emulation。
    """
    fg = _capture_foreground()
    try:
        page = ctx.new_page()
        try:
            page.set_viewport_size({"width": 414, "height": 896})
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
        except Exception:
            pass
    finally:
        _restore_foreground(fg)
    return page


def assert_mobile_origin(pg):
    """★ 硬檢查: 採集分頁必須在手機端 m.1688.com

    2026-11-06 實測事故: Chrome 若用「桌機 UA」, m.1688.com 會被導到
    www.1688.com (桌機版) → 採集器就默默跑在桌機版的 origin 上,
    與使用者歷來的「play打開1688手機端」流程不一致, 且 mtop 指紋來源不對。
    → 這裡寧可中止, 也不要默默跑錯版本。
    """
    try:
        u = pg.url or ""
    except Exception:
        return True
    if "m.1688.com" in u:
        return True
    log(f"  [STOP] 採集分頁不在手機端 (m.1688.com): {u[:90]}")
    log("         → Chrome 必須用「手機 UA」啟動, 否則 m.1688.com 會被導到桌機版")
    log("         → 參考: scripts/start_session.py (請用 --user-agent 手機 UA 的版本)")
    return False


def ensure_page(ctx):
    """確保拿到「專用採集分頁」

    own_tab=True (2026-11-04 新增): 只用「自己開的分頁」, 絕不碰使用者的分頁
      原因 (實測): 使用者在同一個瀏覽器操作 1688 時, 分頁被導走 →
            execution context 被銷毀 → page.evaluate 懸空 300 秒 → 看門狗殺人
      同時, 兩個採集器共用同一個分頁會互相導頁 → 風控風險 ×2
    own_tab=False (舊品類預設): 沿用「重用任何帶 _mc=1 的分頁」的原行為
    """
    global page
    try:
        if page is not None and not page.is_closed():
            return page
    except Exception:
        pass
    if OWN_TAB:
        page = new_background_page(ctx, COLLECT_URL)
        try:
            page.wait_for_timeout(500)
        except Exception:
            pass
        wait_page_settled(page)
        enable_focus_emulation(ctx, page)
        if not assert_mobile_origin(page):
            raise SystemExit(4)
        log(f"  [PAGE] 已開背景採集分頁 (own_tab, 手機端, 不碰使用者分頁、不搶焦點): {page.url[:70]}")
        return page
    for pg in ctx.pages:
        if "_mc=1" in (pg.url or ""):
            page = pg
            enable_focus_emulation(ctx, page)
            wait_page_settled(page)
            return page
    page = ctx.new_page()
    try:
        page.set_viewport_size({"width": 414, "height": 896})
        page.goto(COLLECT_URL, wait_until="domcontentloaded", timeout=30000)
        wait_page_settled(page)
    except Exception:
        pass
    enable_focus_emulation(ctx, page)
    log(f"  [PAGE] 已開專用採集分頁: {page.url[:70]}")
    return page


def repair_page(ctx):
    """分頁 context 失效 (被導走 / 被使用者操作) → 重建採集分頁
    2026-11-04: 分頁被導走時 page.evaluate 會懸空, 必須主動重建再重試
    """
    global page
    try:
        if page is not None and not page.is_closed():
            page.goto(COLLECT_URL, wait_until="domcontentloaded", timeout=30000)
            wait_page_settled(page)
            enable_focus_emulation(ctx, page)
            log("  [REPAIR] 已重建採集分頁 (context 失效)")
            return page
    except Exception as e:
        log(f"  [REPAIR] 重建失敗: {str(e)[:80]}")
    try:
        if page is not None and not page.is_closed():
            page.close()
    except Exception:
        pass
    page = new_background_page(ctx, COLLECT_URL)
    try:
        wait_page_settled(page)
    except Exception:
        pass
    enable_focus_emulation(ctx, page)
    log("  [REPAIR] 已開新採集分頁")
    return page


def parse_jsonp(text):
    i = text.find("(")
    j = text.rfind(")")
    if i < 0 or j < 0:
        return None
    try:
        return json.loads(text[i + 1:j])
    except Exception:
        return None


def _to_float(v):
    try:
        return float(str(v).replace("%", "").replace(",", ""))
    except Exception:
        return None


def offer_to_item(o):
    """API offer 物件 -> 內部 item (欄位口徑已與舊 DOM 檔對齊驗證過)"""
    oid = str(o.get("id") or "").strip()
    if not oid:
        return None
    title = (o.get("simpleSubject") or "").strip()
    if not title:
        title = str(o.get("subject") or "").replace("<font color=red>", "").replace("</font>", "").strip()
    price = _to_float(o.get("price"))
    sales = _to_float(o.get("bookedCount"))          # 與舊 DOM 的「成交N笔」同口徑
    if sales is not None:
        sales = int(sales)
    rp = _to_float(o.get("shopRepurchaseRate"))      # 與舊 DOM 復購率% 同口徑
    tp = _to_float(o.get("tpYear"))
    img = (o.get("offerPicUrl") or "").strip()
    if img.startswith("//"):
        img = "https:" + img
    return {
        "offerId": oid,
        "title": title[:120],
        "priceText": f"￥{price:.1f}" if price is not None else "",
        "priceMin": price,
        "priceMax": price,
        "salesText": f"成交{sales}笔" if sales is not None else "",
        "salesNum": sales,
        "repurchaseRate": rp,
        "shopYears": int(tp) if tp is not None else None,
        "mainImage": img,
        "url": "https://detail.1688.com/offer/" + oid + ".html",
        # v3 新增欄位 (舊 DOM 檔拿不到)
        "offerRepurchaseRate": _to_float(o.get("offerRepurchaseRate")),
        "weight": _to_float(o.get("offerWeight")),
        "monthSales": _to_float(o.get("quantitySumMonth")),
        "isNew": str(o.get("superNewProduct")).lower() == "true",
        "companyName": o.get("companyName") or "",
    }


def is_brand(title):
    return any(b in title for b in BRAND_EXCLUDE)


def is_exclude(title):
    if any(kw in title for kw in EXCLUDE_KEYWORDS):
        return True
    # regex 版排除 (2026-11-06): 需要「後接特定字就不算」的判斷
    # 例: 猫粮(?!碗|盆|...) → 「猫粮碗」是碗不是糧, 不能殺
    for rx in EXCLUDE_REGEX:
        if rx.search(title):
            return True
    return False


def is_too_heavy(title, price):
    if isinstance(price, (int, float)) and price >= 80.0:
        return False
    return any(k in title for k in HEAVY_ITEM_KW)


def passes_filters(it):
    """只判斷「價/品牌/敏感/銷量/超重」, 回傳 (通過?, 淘汰原因)
    跨任務去重另外處理 (因為要先看 keep 白名單)"""
    pr = it.get("priceMin")
    if pr is None or pr < PRICE_MIN or pr > PRICE_MAX:
        return False, "price"
    if is_brand(it.get("title", "")):
        return False, "brand"
    if is_exclude(it.get("title", "")):
        return False, "sens"
    sn = it.get("salesNum")
    if sn is None or sn < MIN_SALES:
        return False, "sales"
    # 銷量上限 (可選, 2026-11-04 為「銷量區間」型任務新增; 沒有此設定的品類不受影響)
    if MAX_SALES is not None and sn > MAX_SALES:
        return False, "sales_over"
    if is_too_heavy(it.get("title", ""), pr):
        return False, "heavy"
    return True, ""


def write_abort_signal(reason, last_kw=""):
    ABORT_FILE.write_text(
        f"timestamp={datetime.now().isoformat()}\nreason={reason}\nlast_keyword={last_kw}\n",
        encoding="utf-8")
    log(f"[ABORT] {reason} (最后词: {last_kw})")


# 限流重試上限 (2026-11-09): 退避重試這麼多次還被限 → 整輪停止, 別硬磨
RATE_LIMIT_MAX_HITS = int(os.environ.get("M1688_RATE_LIMIT_MAX_HITS", "3"))


def _save_response_artifact(ctx, tag, kw, body, kind):
    """把「觸發風控/限流的原始回應」存檔 (2026-11-09)

    沒有這個的話, 事後只知道「被擋了」, 不知道是限流還是滑塊挑戰。
    """
    try:
        import v3_artifacts as ART
        d = ART.capture(ctx, tag=tag, kind=kind, note=f"{kind} @ {kw}")
        (d / "response.txt").write_text(body or "", encoding="utf-8", errors="replace")
        return d
    except Exception:
        return None


# ── 風控守門 (2026-11-04) ────────────────────────────────────────────────
# 見 v3_risk_guard.py: 併發鎖 / 輪次預算 / 每日預算 / 產出熔斷
from v3_risk_guard import (  # noqa: E402
    acquire_lock, release_lock, budget_check, note_word, note_round,
    state_summary, yield_circuit, lock_owner,
)


def preflight_risk(ctx, kwset):
    """開跑前先探一次 mtop, 確認帳號沒被風控擋住 (exit 2 = 風控, 3 = 異常)
    2026-11-04: 沒有這道閘門時, 風控狀態下開跑會懸空 300 秒才被看門狗殺掉
    """
    probe = SCRIPT_DIR / "probe_preflight.py"
    if not probe.exists():
        return True, ""
    import subprocess
    try:
        r = subprocess.run([sys.executable, "-u", str(probe), kwset],
                           capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=90)
        out = (r.stdout or "")[-400:]
        if r.returncode == 0:
            log(f"[PREFLIGHT] OK")
            return True, ""
        if r.returncode == 9:
            # 限流 (2026-11-09): 不是滑塊挑戰 → 退避後可再試, 不該整批放棄
            log(f"[PREFLIGHT] 限流 (可退避重試): {out.strip()[:160]}")
            return False, "RATE_LIMITED"
        if r.returncode == 2:
            log(f"[PREFLIGHT] 風控! {out.strip()[:200]}")
            return False, "preflight 偵測到風控 (滑塊挑戰, 需人工處理)"
        log(f"[PREFLIGHT] 異常 exit={r.returncode}: {out.strip()[:200]}")
        return False, f"preflight 異常 (exit={r.returncode})"
    except Exception as e:
        log(f"[PREFLIGHT] 執行失敗: {str(e)[:120]}")
        return True, ""


DROP_LOG = TEMP_DIR / f"v3_{_KWSET}_{TAG}_dropped.txt"


def log_dropped(kw, why, it):
    """被剔除的商品記錄到側檔, 供複審時檢查篩選是否誤殺"""
    try:
        with open(DROP_LOG, "a", encoding="utf-8") as f:
            f.write(f"{why}\t{kw}\t{it.get('priceMin')}\t{it.get('salesNum')}\t{it.get('title')}\n")
    except Exception:
        pass


def download_image(url, save_path):
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as r:
            data = r.read()
        with open(save_path, "wb") as f:
            f.write(data)
        return True
    except Exception:
        return False


def write_excel(items, table_idx, ts_now):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    headers = ["序号", "商品标题", "主图(嵌图)", "价格", "价格min", "价格max",
               "销量", "复购率%", "开店年数", "商品链接", "品牌/标记", "采集时间",
               "商品复购率%", "重量kg", "30天销量", "新品"]
    col_widths = [6, 50, 16, 14, 10, 10, 12, 9, 9, 30, 12, 18, 11, 9, 10, 7]
    for c in range(1, len(headers) + 1):
        ws.cell(1, c, headers[c - 1]).font = Font(bold=True, color="FFFFFF")
        ws.cell(1, c).fill = PatternFill("solid", fgColor="409EFF")
        ws.cell(1, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = col_widths[c - 1]
    ws.row_dimensions[1].height = 32

    img_dir = TEMP_DIR / f"shopee_images_v3_{_KWSET}_{TAG}"
    img_dir.mkdir(exist_ok=True)

    for i, it in enumerate(items, 1):
        sales_num = it.get("salesNum")
        new_mark = "新" if it.get("isNew") else ""
        brand_mark = "品牌" if it.get("is_brand", False) else ""
        explicit_mark = "敏感词" if is_exclude(it.get("title", "")) else ""
        row = [i, it["title"], it.get("mainImage", ""),
               it.get("priceText", ""), it.get("priceMin", ""), it.get("priceMax", ""),
               it.get("salesText", ""),
               it.get("repurchaseRate") if it.get("repurchaseRate") is not None else "",
               it.get("shopYears") if it.get("shopYears") is not None else "",
               it.get("url", ""),
               f"{brand_mark} {explicit_mark}".strip(), ts_now,
               it.get("offerRepurchaseRate") if it.get("offerRepurchaseRate") is not None else "",
               it.get("weight") if it.get("weight") is not None else "",
               it.get("monthSales") if it.get("monthSales") is not None else "",
               new_mark]
        ws.append(row)
        last_row = ws.max_row
        url_cell = ws.cell(last_row, 10)
        if url_cell.value and str(url_cell.value).startswith("http"):
            url_cell.hyperlink = str(url_cell.value)
            url_cell.font = Font(color="0563C1", underline="single")
        if new_mark:
            nc = ws.cell(last_row, 16)
            nc.fill = PatternFill("solid", fgColor="C6EFCE")
            nc.font = Font(bold=True, color="006100")
        if brand_mark or explicit_mark:
            mc = ws.cell(last_row, 11)
            mc.fill = PatternFill("solid", fgColor="FFE7E6")
            mc.font = Font(bold=True, color="CF1322")
        # 註: 批次檔不嵌圖 — 圖片統一在合併階段下載+壓縮+嵌入 (避免重複下載浪費時間)
        ws.row_dimensions[last_row].height = 120

    out = EXCEL_DIR / f"{TABLE_PREFIX}-m1688-{DATE}-bikeparts-{TAG}-表{table_idx:02d}.xlsx"
    wb.save(str(out))
    return out


def save_all_tables(items, ts_now):
    """把全部商品切成每表 100 筆寫出 (崩潰時資料也不會只留在記憶體)"""
    outs = []
    for idx in range(0, len(items), TARGET_PER_TABLE):
        chunk = items[idx:idx + TARGET_PER_TABLE]
        outs.append(write_excel(chunk, idx // TARGET_PER_TABLE + 1, ts_now))
    if not outs:
        outs.append(write_excel(items, 1, ts_now))
    return outs


def main():
    global page
    all_kw = [l.strip() for l in KW_FILE.read_text(encoding="utf-8").split("\n") if l.strip()]
    batch = all_kw[START - 1:END]
    log(f"m.1688 mtop API v3 {TABLE_PREFIX} {DATE} [{TAG}]")
    log(f"  价格 {PRICE_MIN}-{PRICE_MAX}, 销量 >= {MIN_SALES}")
    log(f"  词 {START}-{END} / 共 {len(batch)} 词 (总表 {len(all_kw)} 词)")
    log(f"  pageSize={PAGE_SIZE}, 单词最多 {MAX_PAGES_PER_KW} 页, 达标 {TARGET_PER_KW} 笔即停")
    log("=" * 60)

    # ── 風控守門 0: 併發鎖 (同一帳號只能有一個採集器) ──
    # 2026-11-04: 你與我同時採集 → 風控風險 ×2, 且互相搶分頁 → 請求懸空
    if not acquire_lock(_KWSET, tag=TAG):
        owner = lock_owner() or {}
        log(f"[GUARD] 採集鎖被占用 (pid={owner.get('pid')} tag={owner.get('tag')}) — 中止, 不要硬敲")
        write_abort_signal("採集鎖被占用(有其他採集器在跑)", "")
        log("[GUARD] 請等對方結束或停掉對方後再跑")
        return
    log(f"[GUARD] 取得採集鎖 (pid={os.getpid()})")
    try:
        return _run_collection(all_kw, batch)
    finally:
        release_lock()


def _run_collection(all_kw, batch):
    existing_oids = load_existing_offerids()
    keep_oids = load_keep_offerids()
    start_watchdog()
    start_ts = time.time()
    deadline = start_ts + HARD_TIMEOUT_SEC
    new_items = []
    new_oids = set()
    pool_saved = set()
    seen_run = set()
    low_streak = 0

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx_list = browser.contexts
        ctx = ctx_list[0] if ctx_list else browser.new_context()
        # v3 走 API, 不需要也不該動使用者的分頁 (使用者可能同時在 1688 上搜尋)
        # 2026-11-04: 原本這裡自己挑分頁, 會抓到「使用者正在用的」那個 (帶 _mc=1) →
        #   使用者一操作就導頁 → context 被銷毀 → 請求懸空 300 秒
        #   現在統一走 ensure_page(): own_tab=True 時只用自己的分頁 + focus emulation + 等 URL 穩定
        ensure_page(ctx)
        log(f"  目前分頁數 {len(ctx.pages)} (使用者的分頁不動)")

        token = ""
        for c in ctx.cookies():
            if c["name"] == "_m_h5_tk":
                token = c["value"].split("_")[0]
        if not token:
            # _m_h5_tk 是 session cookie, 閒置久了會消失 → 重新整理首頁即可拿回
            log("[TOKEN] 開頭拿不到 _m_h5_tk, 重新整理採集分頁刷新 ...")
            try:
                ensure_page(ctx).goto(COLLECT_URL, wait_until="domcontentloaded", timeout=30000)
                time.sleep(5)
                for c in ctx.cookies():
                    if c["name"] == "_m_h5_tk":
                        token = c["value"].split("_")[0]
            except Exception as e:
                log(f"[TOKEN] 刷新失敗: {e}")
        if not token:
            log("[FATAL] 拿不到 _m_h5_tk token, 需重新登入")
            return
        log(f"  token: {token[:24]}...")

        # ── 風控守門 1: preflight (帳號是否已被風控擋住) ──
        # ★ 2026-11-09: 限流(RATE_LIMITED)可退避重試 → 自行等待再試, 不整批放棄
        #   (滑塊挑戰 RISK_CONTROL 才需要人處理 → 直接停)
        pf_tries = int(os.environ.get("M1688_PREFLIGHT_TRIES", "4"))
        for _pf_i in range(pf_tries):
            ok_pf, why_pf = preflight_risk(ctx, _KWSET)
            if ok_pf:
                break
            if why_pf == "RATE_LIMITED" and _pf_i < pf_tries - 1:
                _w = random.uniform(75, 120)
                log(f"[GUARD] 限流 → 等待 {_w:.0f} 秒後重試 ({_pf_i+2}/{pf_tries})")
                time.sleep(_w)
                continue
            log(f"[GUARD] {why_pf}")
            log("[GUARD] 停止 (不硬敲); 依 SOP 冷卻 40-50 分鐘並由使用者手動暖機後再跑")
            write_abort_signal(f"preflight 失敗: {why_pf}", "")
            return

        done_kw = load_done_keywords()
        if done_kw:
            log(f"  [RESUME] 已完成詞 {len(done_kw)} 個, 自動跳過")
        log(f"  [GUARD] {state_summary(_KWSET)}")

        recent_yields = []
        recent_raws = []
        round_used = 0

        for i, kw in enumerate(batch, 1):
            if time.time() > deadline:
                log("[TIMEOUT] 达到硬切时间, 退出主循环")
                break
            # ── 風控守門: 輪次/每日預算 ──
            ok_budget, why_budget = budget_check(_KWSET, round_used)
            if not ok_budget:
                log(f"[GUARD] {why_budget}")
                log("[GUARD] 優雅停止 (已收資料都會存檔, 用續跑記錄可接著跑)")
                break
            elapsed = int(time.time() - start_ts)
            log(f"[{START + i - 1}/{len(all_kw)}] {kw} (已 {elapsed}s)")
            if kw in done_kw:
                log("  [SKIP] 已完成過, 跳過 (省一次請求)")
                continue
            round_used += 1

            # ── 抓取單詞所有頁 (含重試) ──
            kw_items = []
            kw_seen = set()
            n_req = 0
            n_ok = 0                # 成功拿到 API 回應的頁數 (0 = 帳號降級/接口故障)
            hard_stop = None       # 真風控 → 必須中止
            soft_err = None        # 伺服器偶發逾時 → 記錄後繼續
            rate_hits = 0          # 本詞的限流命中次數 (2026-11-09)

            # ── 擬人流量 (2026-11-06): 真人搜尋會先發 suggest/hotkey/statistics ──
            # 目的: 讓請求圖不再是「只有 seo.offer.get 的孤島」
            # M1688_HUMAN_TRAFFIC=0 可關閉 (除錯用)
            if os.environ.get("M1688_HUMAN_TRAFFIC", "0").strip() != "0":
                try:
                    pg_now = ensure_page(ctx)
                    for _label, _u in build_human_traffic_urls(token, kw):
                        jitter(random.uniform(0.25, 0.9))
                        fetch_mtop(pg_now, _u)
                        n_req += 1
                except Exception as e:
                    log(f"  [WARN] 擬人流量前置請求失敗 (不影響主流程): {str(e)[:70]}")

            for page_no in range(1, MAX_PAGES_PER_KW + 1):
                got = False
                for attempt in range(1, RETRY_PER_PAGE + 1):
                    url = build_mtop_url(token, kw, page_no, callback=f"mtopjsonp{page_no}")
                    pg_now = ensure_page(ctx)
                    r = fetch_mtop(pg_now, url)
                    n_req += 1
                    # 真風控優先判斷 (2026-11-04 新增): 請求途中被導到驗證頁 → 立刻中止
                    if r.get("status") == -2:
                        hard_stop = "風控攔截(請求途中被導至驗證頁)"
                        log(f"  [STOP] {hard_stop}")
                        break
                    if not r["ok"] or r["status"] != 200:
                        soft_err = f"HTTP {r['status']}"
                        log(f"  [RETRY] page{page_no} HTTP {r['status']} "
                            f"({attempt}/{RETRY_PER_PAGE})")
                        # 2026-11-04: context 被銷毀 (分頁被導走) → 重建分頁再重試
                        if r["text"].startswith("EVAL_ERR"):
                            repair_page(ctx)
                        time.sleep(random.uniform(RETRY_BACKOFF_MIN, RETRY_BACKOFF_MAX))
                        continue
                    txt = r["text"]
                    # ★ 2026-11-09 判定順序修正 (實測踩到):
                    #   1688 的「限流」回應會同時帶 punish 網址 + captcha 字樣:
                    #     ret   = ["FAIL_SYS_USER_VALIDATE","RGV587_ERROR::SM::哎哟喂,被挤爆啦,请稍后重试"]
                    #     data.url = ".../_____tmd_____/punish?...&action=captcha&pureCaptcha="
                    #   → 先用舊的硬標記清單會誤判成「滑塊挑戰」而直接停機;
                    #     但它其實是**限流** → 正確處置是退避 60-90 秒重試。
                    #   ★ 使用者 2026-11-06「滑動驗證一直過不了」的真實原因就是這個:
                    #     那不是解不開的挑戰, 是需要「等」的限流懲罰頁 → 拖滑塊沒有用。
                    _kind, _ind = PS.classify_from_body(txt, "")
                    if _kind == PS.RATE_LIMITED:
                        _dec = PS.recovery_decision(_kind)
                        rate_hits += 1
                        log(f"  [RATE] 限流 ({','.join(_ind)}) → 退避 {_dec['cooldown_sec']}s 後重試"
                            f" (第 {rate_hits} 次)")
                        _save_response_artifact(ctx, TAG, kw, txt, "rate_limited")
                        if rate_hits >= RATE_LIMIT_MAX_HITS:
                            hard_stop = "限流持續 (RGV587/被挤爆啦) — 退避重試仍被限, 整輪停止"
                            break
                        time.sleep(random.uniform(_dec["cooldown_sec"],
                                                 _dec["cooldown_sec"] + 30))
                        continue
                    if _kind == PS.RISK_CHALLENGE:
                        hard_stop = "風控攔截(回應含明確驗證徵兆)"
                        _save_response_artifact(ctx, TAG, kw, txt, "risk_challenge")
                        break
                    if any(t in txt for t in ("_____tmd_____", "sec_verify", "滑块验证",
                                             "captcha", "verify-slider", "Access Denied")):
                        hard_stop = "風控攔截(回應含明確驗證徵兆)"
                        _save_response_artifact(ctx, TAG, kw, txt, "hard_marker")
                        break
                    if "访问被拒绝" in txt or "訪問被拒絕" in txt:
                        # 可能是真攔截, 也可能是誤判 → 用「頁面是否被導到 punish」二次確認
                        try:
                            now_url = ensure_page(ctx).url or ""
                        except Exception:
                            now_url = ""
                        if any(t in now_url for t in ("_____tmd_____", "punish", "sec_verify")):
                            hard_stop = "訪問被拒絕/風控攔截"
                            break
                        log("  [WARN] 回應含「訪問被拒絕」但頁面未被導走 → 不中止, 視為誤判重試")
                        soft_err = "訪問被拒絕(疑似誤判)"
                        break
                    j = parse_jsonp(txt)
                    if not j:
                        soft_err = "JSONP 解析失敗"
                        log(f"  [RETRY] page{page_no} JSONP 解析失敗 "
                            f"({attempt}/{RETRY_PER_PAGE})")
                        time.sleep(random.uniform(RETRY_BACKOFF_MIN, RETRY_BACKOFF_MAX))
                        continue
                    ret = str(j.get("ret") or "")
                    if ret and "SUCCESS" not in ret:
                        # token 過期 → 刷新後重試
                        if "TOKEN" in ret.upper() or "EXPIRE" in ret.upper():
                            try:
                                ensure_page(ctx).goto(COLLECT_URL,
                                                      wait_until="domcontentloaded", timeout=20000)
                                time.sleep(3)
                                for c in ctx.cookies():
                                    if c["name"] == "_m_h5_tk":
                                        token = c["value"].split("_")[0]
                                log(f"  [TOKEN] 已刷新 token: {token[:20]}...")
                            except Exception:
                                pass
                            time.sleep(random.uniform(RETRY_BACKOFF_MIN, RETRY_BACKOFF_MAX))
                            continue
                        # 伺服器暫時性逾時 (FAIL_SYS_HSF_ASYNC_TIMEOUT 等) → 重試, 不中止
                        soft_err = ret[:110]
                        log(f"  [RETRY] page{page_no} {ret[:80]} "
                            f"({attempt}/{RETRY_PER_PAGE})")
                        time.sleep(random.uniform(RETRY_BACKOFF_MIN, RETRY_BACKOFF_MAX))
                        continue
                    got = True
                    soft_err = None
                    n_ok += 1
                    break
                if hard_stop:
                    break
                if not got:
                    break

                d = j.get("data") or {}
                offs = d.get("offers") or []
                for o in offs:
                    it = offer_to_item(o)
                    if it and it["offerId"] not in kw_seen:
                        kw_seen.add(it["offerId"])
                        kw_items.append(it)
                n_pass = sum(1 for it in kw_items if passes_filters(it)[0])
                if n_pass >= TARGET_PER_KW:
                    log(f"  [够数即停] page{page_no} 已有 {n_pass} 笔通过筛选, 提早停")
                    break
                if not offs:
                    break
                if page_no < MAX_PAGES_PER_KW:
                    jitter(random.uniform(REQ_GAP_MIN, REQ_GAP_MAX))

            if hard_stop:
                log(f"  [STOP] {hard_stop}")
                # ★ 2026-11-09: 中止時留 artifact (截圖+HTML+文字+網路摘要)
                #   過去只有一行 log, 事後無法分辨「風控 / 改版 / 網路 / 分頁被搶」
                try:
                    import v3_artifacts as ART
                    _d = ART.capture(ctx, tag=TAG, kind="hard_stop", note=f"{hard_stop} @ {kw}")
                    log(f"  [ARTIFACT] {str(_d)}")
                except Exception as _e:
                    log(f"  [ARTIFACT] 留證失敗: {str(_e)[:70]}")
                write_abort_signal(f"API: {hard_stop}", kw)
                break
            if soft_err:
                log(f"  [WARN] 最後一次請求失敗: {soft_err}")
            # 搜尋後真人還會補的請求 (擬人流量)
            if os.environ.get("M1688_HUMAN_TRAFFIC", "0").strip() != "0":
                try:
                    pg_now = ensure_page(ctx)
                    for _label, _u in build_human_traffic_after(token, kw):
                        jitter(random.uniform(0.2, 0.7))
                        fetch_mtop(pg_now, _u)
                        n_req += 1
                except Exception:
                    pass
            log(f"  -> {len(kw_items)} 筆 ({n_req} 次請求, {len(kw_seen)} 唯一)")

            if n_pass < 10:
                # 壞詞 (本身沒商品 / 全被敏感規則擋掉) ≠ 帳號降級, 不列入降級判定
                log(f"  [低產出] {kw} 僅 {n_pass} 筆通過 (原始 {len(kw_items)} 筆) — 不算降級")
            if n_ok == 0:
                # 這個詞連一次 API 都沒成功 → 才是真正的降級/接口故障
                low_streak += 1
                log(f"  [接口全失敗] 連續 {low_streak} 個詞 (原始 {len(kw_items)} 筆)")
            else:
                low_streak = 0
            if low_streak >= 2:
                log(f"[STOP] 連續 {low_streak} 個詞 API 全失敗, 判定帳號降級, 立即停止")
                write_abort_signal("連續 API 全失敗(帳號降級)", kw)
                break

            # ── 風控守門: 產出量熔斷 (排序限流偵測) ──
            # 排序限流的特徵: 每詞產出量持續下滑 (11-22 → 3-4 筆/詞)
            # 但要排除「關鍵詞本身稀疏」→ 原始筆數也塌了就不是帳號限流
            recent_yields.append(n_pass)
            recent_raws.append(len(kw_items))
            tripped, why_y = yield_circuit(_KWSET, recent_yields, recent_raws)
            if tripped:
                log(f"[GUARD] {why_y}")
                write_abort_signal(f"產出量熔斷(排序限流): {why_y}", kw)
                break

            # ── 篩選 ──
            c_price = c_brand = c_sales = c_sales_over = c_dedup = c_exp = c_heavy = c_inrun = 0
            n_pass = 0
            for it in kw_items:
                ok, why = passes_filters(it)
                if not ok:
                    if why == "price":
                        continue
                    if why == "brand":
                        c_brand += 1
                    elif why == "sens":
                        c_exp += 1
                    elif why == "sales":
                        c_sales += 1
                    elif why == "sales_over":
                        c_sales_over += 1
                    elif why == "heavy":
                        c_heavy += 1
                    log_dropped(kw, why, it)
                    continue
                c_price += 1
                n_pass += 1
                oid = it["offerId"]
                if oid in seen_run:
                    c_inrun += 1
                    continue
                if oid in existing_oids and oid not in keep_oids:
                    c_dedup += 1
                    continue
                seen_run.add(oid)
                it["is_brand"] = False
                new_items.append(it)
                new_oids.add(oid)
            log(f"  抓 {len(kw_items)} → 價帶內 {c_price} → 品牌 {c_brand} → 敏感 {c_exp} → "
                f"銷量低 {c_sales} → 銷量超 {c_sales_over} → 超重 {c_heavy} → "
                f"跨任務 {c_dedup} → 本輪重複 {c_inrun} → "
                f"累计 {len(new_items)}")

            mark_done(kw)
            # ── 風控守門: 記錄本詞用量 (輪次/每日預算用) ──
            note_word(_KWSET, requests=n_req, got=len(new_items))

            if i % AUTO_SAVE_EVERY == 0 and new_items:
                outs = save_all_tables(new_items, datetime.now().strftime("%Y-%m-%d %H:%M"))
                log(f"[AUTO-SAVE] {len(outs)} 個表 (已收 {len(new_items)} 条)")
                # 同步寫入去重池 (中途崩潰也不會讓已收商品變成「可重複採集」)
                pend = new_oids - pool_saved
                if pend:
                    save_oid_pool(pend)
                    pool_saved |= pend

            if i < len(batch):
                gap = random.uniform(KW_GAP_MIN, KW_GAP_MAX)
                extra = ""
                if random.random() < LONG_PAUSE_CHANCE:
                    lp = random.uniform(LONG_PAUSE_SEC_MIN, LONG_PAUSE_SEC_MAX)
                    gap += lp
                    extra = f" + 长休 {lp:.0f}s"
                if i % random.randint(COOL_DOWN_EVERY_MIN, COOL_DOWN_EVERY_MAX) == 0:
                    cd = random.uniform(COOL_DOWN_SEC_MIN, COOL_DOWN_SEC_MAX)
                    gap += cd
                    extra += f" + 冷却 {cd:.0f}s"
                log(f"[节奏] 停顿 {gap:.0f}s{extra} -> 下一词")
                touch_hit()          # 休息期間不算卡死
                jitter(gap)
                touch_hit()

    ts_now = datetime.now().strftime("%Y-%m-%d %H:%M")
    outs = save_all_tables(new_items, ts_now)

    if new_oids:
        pend = new_oids - pool_saved
        if pend:
            save_oid_pool(pend)
    log(f"\n[完成] {TAG}: {len(new_items)} 条, 新 offerId {len(new_oids)}")
    for o in outs:
        log(f"  -> {o}")
    log(f"  用时 {int(time.time() - start_ts)}s")


if __name__ == "__main__":
    main()
