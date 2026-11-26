# 2799漏勺（repo: 2799-loushao）— 1688 選品採集技能（跨裝置可攜）

> 分組：shopee小巧思 ｜ 倉庫名用拼音是 GitHub 限制（repo 名不接受中文）
> 改名目的：避免他人直接搜尋「1688 選品 skill」就取得可用的腳本

> 這是把 DSH 技能 `m1688-selection-ops` 打包成**獨立可攜**的版本。
> 換一台電腦時：下載本 repo → 執行 `install.ps1` → 把技能資料夾放進你的 client 的 skills 目錄。

---

## 一、這個技能在做什麼

在 **1688（m.1688.com / s.1688.com）** 上依「關鍵字 + 價格帶 + 銷量區間 + 品牌排除」篩選商品，
輸出帶主圖的 Excel 選品總表。核心特性：

| 特性 | 說明 |
|---|---|
| **不靠 DOM 捲動** | 直接打 mtop 內部搜尋 API，單詞 3–5 秒 |
| **PC 第二通道** | `s.1688.com` 頁面驅動（每頁 120 筆），桌面 UA + GBK 關鍵字 |
| **風控四道閘門** | 併發鎖 / 輪次預算 / 每日預算 / 產出熔斷（程式強制，不靠人自律）|
| **五態頁面分類** | normal / not_logged_in / risk_challenge / rate_limited / site_changed |
| **驗收六項** | 圖片逐列位元組相符、序號連續、價量零越界、列高、空連結、品牌殘留 |

---

## 二、換裝置安裝（3 步）

### 1. 下載

```powershell
git clone <這個 repo 的網址>
# 或直接下載 zip 再解壓
```

### 2. 安裝技能

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

`install.ps1` 會：
- 把 `m1688-selection-ops/` 複製到目標 client 的 `skills/` 目錄
  （預設 `C:\Users\<你>\.dsh\profiles\<profile>\node_modules\dsh-minimax-workflows\skills\`）
- 列出 `scripts/` 的核心腳本與它們的依賴

### 3. 環境自檢

```powershell
C:\Python314\python.exe -u scripts\check_env.py
```

---

## 三、★ 開 Chrome 的 debug 模式（採集前必做）

採集器是透過 **CDP 接管 Chrome**，不是自己開一個無痕瀏覽器 —— 這樣才能帶著你的 1688 登入態。

### 3.1 手機通道（9223，主力，UA 必須是手機版）

> ⚠ **UA 一定要是手機版**。用桌面 UA 時 `m.1688.com` 會自動導向 `www.1688.com`（桌面版），
> 採集器的 `assert_mobile_origin` 檢查會中止（exit code 4）。

```powershell
$chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$argline = '--remote-debugging-port=9223 --user-data-dir="C:\Users\<你>\AppData\Local\Google\Chrome\User Data_1688" --user-agent="Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36" --no-first-run --no-default-browser-check --disable-features=Translate,OptimizationGuideModelDownloading --window-size=430,950'
Start-Process -FilePath $chrome -ArgumentList $argline
```

### 3.2 PC 通道（9226，每頁 120 筆、GBK 關鍵字）

```powershell
$chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$argline = '--remote-debugging-port=9226 --user-data-dir="C:\Users\<你>\AppData\Local\Google\Chrome\User Data_1688_pc" --no-first-run --no-default-browser-check --disable-features=Translate --window-size=1400,900 "https://s.1688.com/"'
Start-Process -FilePath $chrome -ArgumentList $argline
```

### 3.3 三個實例一次開（建議用這個）

```powershell
C:\Python314\python.exe -u scripts\ensure_chrome.py 9223   # 手機版
C:\Python314\python.exe -u scripts\ensure_chrome.py 9226   # 桌機版
```

`ensure_chrome.py` 會先檢查埠是否活著，**活著就不重開**（重開＝新 session＝額外風控層）。

### 3.4 只關 9223 實例（不動你的一般 Chrome）

```powershell
Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
  Where-Object { $_.CommandLine -like '*remote-debugging-port=9223*' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
```

---

## 四、換裝置後需要自己設定的東西

| 項目 | 說明 |
|---|---|
| **1688 帳號** | 需在 9223／9226 的 Chrome 視窗**手動登入**同一個 1688 帳號（帳號信譽＝風控分數）|
| **profile 目錄** | `User Data_1688` / `User Data_1688_pc` 需自建；登入態不會跟著 repo 走 |
| **採集預算** | `v3_risk_guard.py` 預設：單輪 20 詞 / 每日 60 詞 / 產出熔斷 3 筆每詞 |
| **路徑** | 腳本裡的路徑寫死 `E:\1shopee\...`，換槽位要改 `v3_config.py` 的 `EXCEL` / `TMP` |
| **Python** | 需 Python 3.11+ 且有 `playwright`、`openpyxl`、`Pillow`（`pip install playwright openpyxl pillow`；`playwright install chromium`）|

---

## 五、目錄結構

```
m1688-selection-ops/
├── SKILL.md                     ← 現行作業規則（硬約束、風控四步、欄位口徑、驗收六項、踩坑表）
├── references/
│   └── w1-1688-selection.md     ← 完整規格 29KB（逐條細節、逐品類歷史）
└── scripts/                     ← 可直接跑的核心腳本
    ├── ensure_chrome.py         ← ★ 啟動 CDP 實例（手機 UA / 桌面 UA）
    ├── collect_m1688_v3_api.py  ← 主採集器（手機通道，mtop API v3）
    ├── collect_m1688_pc.py      ← PC 通道採集器
    ├── v3_config.py             ← 品類設定（價帶 / 銷量 / 規則）
    ├── v3_rules.py              ← 共用規則引擎（判定順序）
    ├── v3_risk_guard.py         ← 風控守門（四道閘門）
    ├── merge_v3.py              ← 管線①合併去重 + 下載主圖
    ├── finalize_v3.py           ← 管線②複審篩選 + 嵌圖
    ├── verify_v3.py             ← 管線③驗收
    ├── final_check_cat.py       ← 最終檢查（品牌/IP 殘留掃描）
    └── check_env.py             ← 環境自檢
```

---

## 六、常用指令

```powershell
cd scripts

# ① 採集（手機通道）: <起> <迄> <tag> <詞表名>
C:\Python314\python.exe -u collect_m1688_v3_api.py 1 20 a1 keywords_pet_feed.txt

# ② 採集（PC 通道）: <起> <迄> <tag> <詞表名> <埠>
C:\Python314\python.exe -u collect_m1688_pc.py 1 20 a1 keywords_pet_feed.txt 9226

# ③ 管線
C:\Python314\python.exe -u merge_v3.py <品類>
C:\Python314\python.exe -u finalize_v3.py <品類>
C:\Python314\python.exe -u verify_v3.py <品類>
C:\Python314\python.exe -u final_check_cat.py <品類>
```

**`tag` 必須是新值**（會覆蓋同名批次檔）。

---

## 七、踩過的坑（精選，完整版在 SKILL.md）

| 坑 | 處理 |
|---|---|
| 桌面 UA → `m.1688.com` 轉桌面版 → 採集中止 | 必須手機 UA 開 9223 |
| `afterPrice` 在 PC 回應是**運費**不是銷量 | 銷量以 `bookedCount` 為主 |
| 「书桌／办公桌」是**用途修飾語**不是品類主詞 | 不要放進 `exclude`，交給 `must_words` |
| `keep_rules` 是**純 pattern 列表**、`deny_rules` 才是 tuple 組 | 格式寫錯會拋 TypeError |
| GBK vs UTF-8 關鍵字 | PC 通道必須 GBK |
| 限流 ≠ 滑塊 | 限流退避即可；滑塊要人拖 |
| 採集器搶使用者的分頁 | 用專用分頁 + 開完還 OS 焦點 + 用完關分頁 |

---

## 八、授權

技能內容與腳本隨意取用、修改、二次發布。抓到的是 1688 的公開商品資料；
請遵守 1688 的 robots 與服務條款、**不要**用高併發大量抓取（本技能的守門就是為了這點）。
