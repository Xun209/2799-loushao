# -*- coding: utf-8 -*-
"""v3_rules.py — v3 複審規則的統一套用邏輯 (所有 v3 品類共用)

判定順序 (先命中先決定):
  1. brand_rules    → 品牌 / 仿牌 / IP, 直接刪
  1b. brand_ctx_rules → 車廠品牌, 僅在「無相容性文案」時刪 (見下)
  2. keep_rules     → 保護清單, 命中就整條放行 (跳過 must_words 與 nonfit_rules)
  3. deny_rules     → 否決清單, 可覆蓋 keep_rules 的保護 (放行後仍可精準殺掉)
  4. must_words     → 正面條件, 標題不含本類關鍵字就刪
  5. nonfit_rules   → 非本類 (服飾 / 家具 / 食物 / 電器 …) — 一律套用
  6. nonfit_ctx_rules → 場景型非本類 (汽車 / 商用 / 戶外 …), 有居家語境就豁免

語境豁免 (2026-10-28 垃圾桶品類新增):
  賣家會把可愛造型款寫成場景關鍵字 (「汽車造型垃圾桶家用客廳」「垃圾桶兼收納桶」),
  純字串規則會把本類最好的品殺掉 → nonfit_ctx_rules 只在「沒有居家語境」時才套用。
  詳見 is_ctx_allowed()。

車廠品牌的相容性豁免 (2026-10-21 使用者決定「車廠名是主詞就刪、寫適用XX就留」):
  標題含相容性文案 (适用/適用/適合/兼容/通用/專用/適配/安裝/改裝 …) 時,
  車廠名只是相容性說明, 不算品牌 → brand_ctx_rules 整組跳過。
  實例: 「適用雅马哈本田尾箱」留; 「雅马哈摩托车尾箱」刪。

車廠品牌的「專屬改裝件」豁免 (2026-10-30 使用者決定):
  上條規則把「車廠專屬改裝件」也一起殺了 (機車尾箱一次刪 25 哈雷 / 2 雅马哈 / 2 Vespa / 1 川崎),
  但那正是本類要賣的貨 (「哈雷 883 改装后货架」「雅马哈 Tenere 700 摩托车配件后尾箱架」)。
  使用者定案: 進口車廠的專屬改裝件要保留 → 由品類自己的 brand_ctx_keep 規則表判斷
  (車廠名 + 本類品項字 + 機車語境), 命中就整組跳過 brand_ctx_rules。實作見 moto_rules.py。

為什麼要抽成共用模組:
  finalize_v3.py / verify_v3.py / test_rules.py 各自一份 hit_reason 會漂移。
  2026-10-20 cup 品類就發現 verify_v3.py 少了 keep_rules, 會把
  「被保護而保留」的品項誤判成不該存在, 造成驗證假失敗。
"""
import re


# ── 平台禁賣品 (所有品類一律擋, 不需要各品類自己寫) ─────────────────────
# Shopee 禁賣品: 護目鏡 / 防護眼鏡 / 滑雪鏡 / 游泳鏡 等任何形式的「眼睛防護用品」。
# 使用者 2026-08-27 定案: 任何 1688 選品腳本都不得包含護目鏡相關詞
# (2026-10-28 實測發現: 攀岩類搜尋會撈到「攀岩保護眼鏡」這類商品)
BANNED = [
    "护目镜", "護目鏡", "护目", "護目", "防护眼镜", "防護眼鏡",
    "保护眼镜", "保護眼鏡", "运动眼镜", "運動眼鏡",
    "眼镜", "眼鏡", "泳镜", "泳鏡", "滑雪镜", "滑雪鏡",
    "潜水镜", "潛水鏡", "风镜", "風鏡", "防风镜", "防風鏡",
]


def hit_banned(t):
    """命中禁賣品 → 回傳字串; 沒中回傳 None"""
    for w in BANNED:
        if w in t:
            return f"禁賣品({w})"
    if "goggles" in t.lower():
        return "禁賣品(goggles)"
    return None


def is_compat_title(t, C):
    """標題是否含相容性文案 (適用XX / XX專用 / 通用 …)"""
    ca = C.get("compat_allow")
    return bool(ca) and bool(re.search(ca, t))


def is_brand_ctx_keep(t, C):
    """車廠品牌豁免 (2026-10-30 使用者決定): 命中 → 車廠名不算品牌, 整組跳過 brand_ctx_rules

    為什麼需要:
      2026-10-21 的規則是「車廠名是主詞就刪」, 結果把「車廠專屬改裝件」也一起殺了
      (機車尾箱一次刪 25 哈雷 / 2 雅马哈 / 2 Vespa / 1 川崎), 但那正是本類要賣的貨:
        「哈雷 883 1200 黛娜大滑翔 改装后备箱支架 后货架」
        「跨境雅马哈Tenere 700 新款摩托车配件后货架行李架后尾箱架」
      使用者 2026-10-30 定案: 進口車廠的專屬改裝件要保留。

    判斷由品類自己的 brand_ctx_keep 規則表提供 (見 moto_rules.py):
      車廠名 + 本類品項字 + 機車語境 → 豁免
      (品項字 + 機車語境兩個條件是防誤留: 模型玩具「Maisto 1:18 哈雷…合金摩托模型」
       沒有尾箱/靠背字 → 不豁免; 汽車專屬件沒有機車語境 → 不豁免)
    """
    for pat, _name in C.get("brand_ctx_keep", []):
        if re.search(pat, t):
            return True
    return False


def is_ctx_allowed(t, C):
    """語境豁免: 標題同時含「本類品項字」與「居家/使用場景字」→ 場景型非本類規則不適用

    為什麼需要 (2026-10-28 垃圾桶實測):
      賣家會把「造型可愛款」寫成場景型關鍵字, 例如
        「卡通垃圾桶家用带盖…移动收纳桶」   → 垃圾桶兼收納桶, 純字串規則會誤殺
        「创意小汽车垃圾桶高颜值带盖家用夹缝可移动巴士垃圾筒」 → 汽車造型垃圾桶
        「高脚酒杯垃圾桶…复古…中古风纸篓」   → 中古风是風格不是二手
      這些正好是本類要的高顏值/可愛/個性款, 不能只靠字串擋。

    規則:
      ctx_product  = 本類品項字 (垃圾桶/垃圾筒/紙簍…)
      ctx_allow    = 居家/使用場景字 (家用/客廳/臥室/廚房…)
      兩者同時命中 → nonfit_ctx_rules 整組跳過
    """
    prod = C.get("ctx_product")
    allow = C.get("ctx_allow")
    if not prod or not allow:
        return False
    if not re.search(prod, t):
        return False
    return bool(re.search(allow, t))


def hit_reason(C, t):
    """回傳刪除原因字串; None 代表保留"""
    # 0. 平台禁賣品 — 所有品類一律擋 (護目鏡/防護眼鏡/滑雪鏡/游泳鏡…)
    banned = hit_banned(t)
    if banned:
        return banned
    for pat, name in C.get("brand_rules", []):
        if re.search(pat, t):
            return name
    # 車廠品牌: 有相容性文案、或命中「車廠專屬改裝件」豁免時就不算品牌 (整組跳過)
    if not is_compat_title(t, C) and not is_brand_ctx_keep(t, C):
        for pat, name in C.get("brand_ctx_rules", []):
            if re.search(pat, t):
                return name
    for pat in C.get("keep_rules", []):
        if re.search(pat, t):
            # 保護清單命中, 但 deny_rules 仍可覆蓋 (放行得太寬時用)
            for dpat, dname in C.get("deny_rules", []):
                if re.search(dpat, t):
                    return dname
            return None
    must = C.get("must_words")
    if must and not re.search(must, t):
        return "非本類(無本類關鍵字)"
    for rule in C.get("nonfit_rules", []):
        if isinstance(rule, tuple):
            pat, name = rule
            if re.search(pat, t):
                return name
        else:
            if rule in t:
                return rule
    # 場景型規則: 有居家語境就豁免 (汽車造型/雙功能垃圾桶要留)
    if not is_ctx_allowed(t, C):
        for rule in C.get("nonfit_ctx_rules", []):
            if isinstance(rule, tuple):
                pat, name = rule
                if re.search(pat, t):
                    return name
            else:
                if rule in t:
                    return rule
    return None


def parse_sales(v):
    """銷量欄位「成交X笔」/「成交X万」都要處理萬單位"""
    s = str(v or "").replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*万", s)
    if m:
        return float(m.group(1)) * 10000
    m = re.search(r"\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None
