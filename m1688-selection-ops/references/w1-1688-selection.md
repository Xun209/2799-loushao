# W1 1688 選品採集 — 完整蒸餾規格

## 2. W1 — 1688 選品採集（m.1688.com）

### 2.1 Goal & scope

- 來源平台：`m.1688.com`（手機版搜尋頁）；賣場為 Shopee Taiwan（**YIChu 意厨／柏森蘭／PLT-普籃徒**，以 HHH ERP 管理）。
- 目標：依「品類 → 關鍵詞清單 → 價格區間 → 銷量門檻 → 品牌界線 → 排除詞」採集商品，輸出帶主圖的 Excel 選品總表供人工審查（使用者手動加 `1` 前綴 = 已審查）。
- 兩條技術路線（**v3 為現行，v2 已停止使用**）：
  - v2（DOM 捲動採集）：`collect_m1688_v2_*.py`，已保留但不再跑。
  - v3（mtop JSON API）：`collect_m1688_v3_api.py`，現行。
- 跨任務累積量：2026-08 至 2026-10 共交付 12+ 個品類總表；跨任務去重池 offerId 由 3,891 → 13,153。

### 2.2 Hard constraints / user preferences（硬約束）

**帳號與環境**
- 1688 **單帳號**：`tb966392038`（userid `2206668197833`）——**不能換號**。
- CDP 固定 `http://127.0.0.1:9223`，profile 固定 `User Data_1688`。
- Python `C:\Python314\python.exe`；Windows PowerShell。
- **不重啟 Chrome 實例**（重啟＝新 TCP 連線＋新 session，觸發「異地／新設備」額外風控層）。
- **不同時開兩個 Chrome 實例**（會崩潰且 session 被關聯）。（例外：2026-10-21 使用者明確允許 1688 專用實例與個人 Chrome 雙實例並行。）
- **不用診斷腳本測配額**（診斷本身即搜索請求，會吃掉配額）。
- **不做指紋偽造**（stealth／Canvas 加噪）、**不用代理 IP 池**、**不破解滑塊**——僅採「換 session」與「降低請求量、貼合平台節奏」的合規手段。
- **只關桌機版分頁**（`www.1688.com`／`work.1688.com`），**不可關掉所有分頁**（會導致 Chrome 自關＋`ECONNREFUSED`）；不可誤關 `login.taobao.com` 分頁。
- 採集必須使用**專用分頁** `https://m.1688.com/?_mc=1`，完全不動使用者自己的分頁（腳本不得含任何 `.close()` 分頁程式碼）。
- 使用者明確同意：可用新 Chrome profile 換 session。

**風控處理**
- 遇 `_____tmd_____`／滑塊驗證／`访问被拒绝`／`sec_verify`／`Access Denied`／`captcha`／`verify-slider`／`滑块验证` → **必須立即停止，不自動跳頁重試、不自動破解**，由使用者手動處理。
- 風控標準四步流程（已確立）：**① 拖滑塊 → ② 等 40–50 分鐘 → ③ 使用者手動搜 3–5 個詞暖機 → ④ 啟動採集**。
- 拖完滑塊後搜索配額僅 **2–3 次搜索請求**；**滾動（同頁 XHR）不消耗搜索配額**。
- 風控是**帳號 + IP 層級**：新開 profile（9224 / `User Data_1688_b`）同帳號同樣被拒；隔夜 16 小時不碰 1688 也未恢復 → 換瀏覽器無效。
- 降級（帳號級限流）與壞詞必須分開判定：**只有「連續 2 個詞 API 呼叫全失敗」才算降級**；不把「0 張卡」直接當壞詞（要先測多個詞）。
- 降級判定以**捲動後商品卡數**為準：首屏 11–12 張 = 降級；捲 5 次到 90+ 才算恢復；部分腳本用 `>40` 張、`MIN_CARDS_ENOUGH` 由 60 改 100。不能只看首屏。
- 加入「**6 秒複讀確認**」避免瞬間性誤判。
- 連續 3 詞 <30 張卡 → 自動停止。

**選品條件**
- 售價區間為**硬約束**（per task）。已出現的區間集合：6-21／6-36／6-46／8-48／9-49／9-99／10-30／10-40／10-80／10-100／10-150／10-200／10-300／20-150／25-105／25-300／30-200／40-240／50-300／80-350。
- 銷量門檻為硬約束（per task）：MIN_SALES 1（凱格爾）／50（戶外）／60（健身）／100（多數）／300（trash2）／500（trash）。
- 品牌策略：可接受**中國中階牌**；排除**國際／進口大牌**與**國內頭部品牌**。品牌排除**必須同時覆蓋簡體與繁體寫法**（例：「禧马诺／禧瑪諾」、「蓝图／藍圖」）——先前只寫繁體導致漏抓 124 筆。
- 禁賣品全域護欄（`BANNED`，`v3_rules.py`）：护目镜／護目鏡／护目／護目／防护眼镜／防護眼鏡／保护眼镜／保護眼鏡／运动眼镜／運動眼鏡／眼镜／眼鏡／泳镜／泳鏡／滑雪镜／滑雪鏡／潜水镜／潛水鏡／风镜／風鏡／防风镜／防風鏡 ＋ `goggles` 小寫比對。
- **刀具不是禁賣品**（使用者 2026-10-20 定案）；本輪只擴充「製茶刀具」（茶刀／茶剪／茶叶刀／茶刀套装）。
- 品類特有排除：攻擊性詞（自衛／破窗／擊打／攻擊／防身／武器／凶器／武術／搏擊，33 個）、18 禁／醫療／治療／擦邊詞（凱格爾 50+、按摩舒緩 50+）、保溫瓶污染、承重／重物、服飾（保留手套襪子）。
- 車廠品牌規則（`brand_ctx_rules` + `compat_allow`）：車廠名是主詞就刪、寫「適用 XX」就留；「改装／安装」**不算**相容字樣。
- 排除詞清單在 PDD 側由使用者明令**不啟用**（見 W3）。

**節奏（使用者要求隨機化、不單一）**
- v2 DOM 節奏：詞間停頓 45–110s；捲動 5–9 次；捲動距離 1800–3200；間隔 2.5–5s；冷卻 90–200s；25% 機率長休 120–260s。
- 最保守版：詞間停頓 80–240s（拆 3–6 段不等長、±45% 浮動）；冷卻 150–380s／每 4–8 詞；長休 200–520s／40% 機率；每 2 詞回首頁；捲動 5–10 次；間隔 3.0–7.5s；`HARD_TIMEOUT_SEC = 220*60`。
- **jitter 修正**：拆段 sleep 全部累加會放大 3–6 倍（80–240s 變 525s）；須**權重歸一化**使總時長精確等於 base（2000 次樣本偏差 0）。修正後效率提升約 5 倍，風控耐受從 3–4 詞升到 12–18 詞。
- 100 詞/任務、100 分鐘硬切；auto-partial 每 20 詞留底；每 8 詞 cool-down 60s；每詞回主頁。

**任務管理**
- detach 執行（`Start-Process`），ETA 只設 **1 個 cron（once）**，不循環（使用者反饋「不斷刷新不合理」）。
- 背景 bash 任務上限 3600 秒 → 採集分批，每批 **< 3300 秒**；寧可多切幾批。
- 檔名 v2/v3 為續跑編號，**不修改舊腳本**（複製模板改 hardcode 派生新腳本）。
- 風控中斷詞重跑（v1b / v1c / v1d / v2b 模式）。

### 2.3 Concrete artifacts

**目錄**
```
E:\1shopee\1688選品工作流\scripts\        # 採集與管線腳本（舊根目錄 E:\shopee\1688選品工作流\scripts\）
E:\1shopee\1688選品工作流\keywords_*.txt  # 詞表
E:\1shopee\excel\                         # 全部 xlsx 產出
C:\Users\夏晨茗\AppData\Local\Temp\v3img_<cat>\   # 圖片快取
C:\Users\夏晨茗\AppData\Local\Temp\1688_done_v3_<kc>.txt  # 續採詞表
C:\Users\夏晨茗\.mavis\sessions\mvs_210db7fd001a44fda722bd8129a2d675\workspace\  # 工作區暫存
```

**核心腳本（v3）**
- `collect_m1688_v3_api.py`（採集；52620 bytes 級；簽名／分頁／排序／價格帶／重試／AbortController 逾時／專用分頁／續採詞表／去重池同步寫入／掉品記錄）
- `v3_config.py`（`CATS` 品類設定）→ `v3_rules.py`（共用規則引擎：`brand_rules` / `brand_ctx_rules` / `keep_rules` / `deny_rules` / `must_words` / `nonfit_rules` / `nonfit_ctx_rules` / `ctx_product` / `ctx_allow` / `BANNED` / `compat_allow` / `dedup_against` / `extra_cols`）
- 品類規則檔：`cup_rules.py`、`insul_rules.py`、`moto_rules.py`、`shoe_rules.py`、`trash_rules.py`、`trash2_rules.py`
- 管線：`merge_v3.py`（合併＋並行下圖＋快取）、`finalize_v3.py`（複審＋壓縮嵌入）、`verify_v3.py`（逐列比對）、`final_check_cat.py`、`audit_cat_rules.py`、`check_all_delivered.py`、`check_port.py`、`check_punish.py`
- 輔助：`xlsx_imgfix.py`（`images_by_row` / `xlsx_rowmd5_from_xml`）、`check_env.py`、`open_m1688.py`、`open_1688_only.py`、`open_login.py`、`trigger_punish.py`、`prompt_manual_search.py`、`diag_zero_cards.py`、`start_session.py`、`start_second_session.py`、`shot.py`、`refresh_token.py`、`probe_kw.py`、`fix_missing_img.py`、`dump_all_titles.py`、`show_reason.py`、`audit_drop.py`、`dedup_pool.py`、`warm_cache.py`、`trim_done.py`、`cleanup_batch.py`、`review_titles.py`、`test_rules.py`、`check_kill.py`、`shoes_final.py` + `verify_shoes.py`
- 舊 v2 腳本：`collect_m1688_v2_{tent,fishing_rod_holder,cycling_pants,kids_protection,baseball_bat,waterproof_bag,kegel,relax,grip,fish,stunt,arm,knife,bfly,cycling,gymaux,kidscooter,outdoor,bikeparts}_v*.py`

**詞表**
- `keywords_bike_parts.txt`（原始 200 詞）、`keywords_bike_parts_p1.txt`（100 高價值）、`keywords_bike_parts_p2.txt`（65 備用）
- `keywords_gym_aux.txt`（91 非空詞，實體 97 行）、`keywords_gym_final18.txt`、`keywords_pickleball.txt`（20）
- `keywords_kidscooter.txt`（36）、`keywords_outdoor.txt`（43）
- `keywords_drinkware.txt`（89＝85 杯具茶具咖啡 + 製茶刀具 4）、`keywords_insulated_box.txt`（99）
- `keywords_moto_case.txt`（45）、`keywords_moto_backrest.txt`（71）
- `keywords_dance_shoes.txt`（37）、`keywords_climbing_shoes.txt`（36）、`keywords_water_shoes.txt`（32）
- `keywords_trash_bin.txt`（42）、`keywords_trash2_home_shop.txt`（78）
- `keywords_stunt.txt`（200）、`keywords_arm.txt`（100）、`keywords_knife.txt`（210）、`keywords_bfly.txt`（102）

**去重池**
- `E:\1shopee\excel\_existing_offerids_2026-08-27.txt` —— 單一檔持續累積（不另建日期檔）。累計值：3,891 → 3,877 → 4,982 → 5,121 → 6,041 → 6,427 → 7,827 → 9,241 → 9,666 → 10,779 → 12,206 → 13,153。
- v3 新增 `dedup_mode: "delivered"`：只排除已交付總表 ∪ 本品類既有批次，避免舊任務被篩掉的品項永久封鎖。

**Chrome 啟動（路徑含空格必須用單字串 ArgumentList）**
```powershell
$chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$argline = '--remote-debugging-port=9223 --user-data-dir="C:\Users\夏晨茗\AppData\Local\Google\Chrome\User Data_1688" --no-first-run --no-default-browser-check --disable-features=Translate,OptimizationGuideModelDownloading --window-size=430,950'
Start-Process -FilePath $chrome -ArgumentList $argline
```
**只關 9223 實例（不動使用者一般 Chrome）**
```powershell
Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
  Where-Object { $_.CommandLine -like '*remote-debugging-port=9223*' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
```

### 2.4 Standard procedure

1. **檢查 9223 Chrome 是否活著**；逾 24 小時會死 → 重啟（不得中途中斷使用者的實例）。
2. **確認登入**：`lid=tb966392038`（cookie）。工具：`python 'E:\1shopee\1688選品工作流\scripts\open_1688_only.py' <輸出png>`（導向 `?_mc=1`、印 lid cookie、截圖）。
3. **檢查風控狀態**：`python check_env.py`（開首頁、確認登入、檢查風控、可選關鍵詞測試限流）；滑塊頁必須用 `trigger_punish.py` 保留（`check_env.py` 會把頁面跳回首頁）。若遇風控 → 執行標準四步流程，等使用者回「好了」。
4. **暖機（使用者 2026-10-20 新增的流程步驟）**：`python prompt_manual_search.py 9223` → 開 1688 停在首頁並提醒使用者手動搜 3–5 個詞（建議詞每詞停留 3–5 秒），**等使用者回覆「好了」後才啟動採集**。
5. **建立詞表 + 品類設定**：詞表 `keywords_<cat>.txt`；`v3_config.py` 的 `CATS["<cat>"]`（prefix／sheet／price／min_sales／brand_rules／nonfit_rules／keep_rules／must_words／deny_rules／extra_cols／dedup_against）。
6. **雙向測試篩選規則**（必跑）：應保留案例 0 誤擋、應刪除案例 0 漏擋。腳本 `test_shoe_rules.py` / `test_moto_rules.py` / `test_trash_rules.py` / `test_trash2_rules.py` / `test_rules.py`。
7. **採集**（位置參數 `<start> <end> <tag> <kwset>`）：
   ```
   cd 'E:\1shopee\1688選品工作流\scripts'
   python collect_m1688_v3_api.py 1 85 c1 cup
   ```
   - `<tag>` **必須是新值**，避免覆蓋既有批次檔。
   - 每詞 10–20 秒、請求少 5–10 倍、全程零風控（v3 實測）。
   - 看門狗 thread：300 秒無 API 回應即 `os._exit(3)`；啟動時自動刷新 `_m_h5_tk` token；`AUTO_SAVE_EVERY=2`。
   - 批次檔只存資料（「主图(嵌图)」欄存圖片網址，不嵌圖）。
8. **合併去重**：`python merge_v3.py <cat>` → 合併總表；URL 重複 + 標準化標題重複；**重複商品優先保留有圖片網址的那筆**；並行（8 執行緒）下圖到 Temp 快取。
9. **複審**：`python finalize_v3.py <cat>` → 套品類規則刪除 + 圖片壓縮嵌入 + 產生最終總表。
10. **逐列驗證**：`python verify_v3.py <cat>`（詳見 §2.6）。
11. **交付前複核**：`python final_check_cat.py <cat>`；`python check_all_delivered.py`（檢查所有已交付總表連結無空值）。
12. **稽核與抽樣**：`python audit_cat_rules.py <cat> full`（印各小類刪除明細 + 保留項殘留掃描）；`python test_rules.py <cat> [筆數]`（乾跑，不產檔）；`python check_kill.py <cat> <原因關鍵字>`（看某原因刪掉全部商品，確認是否誤殺）；`falseneg.py <cat> "關鍵字1|關鍵字2"`（列出含本類關鍵字卻被刪的誤殺嫌疑）。
13. **交付**：media 標籤 + 檔案路徑 + 驗證數據 + 刪除明細 + 邊界保留項列表（供使用者回溯調整）。

**v3 API 請求形狀（已實測可用）**
```
端點：h5api.m.1688.com/h5/mtop.1688.wap.seo.offer.get/1.0/
簽名：md5(_m_h5_tk前段 & t & 12574478 & data)
必須在頁面 context 內 fetch（沿用 Chromium TLS 指紋 + cookie，不觸發 document 層配額）
參數：sortType="price_asc"（pop 綜合排序命中率最差：4/60）、priceStart/priceEnd、pageSize=60（200 會被砍回 60）
欄位口徑：銷量=bookedCount（非 saleQuantity）、復購率=shopRepurchaseRate（非 offerRepurchaseRate）、
          價格=price、標題=simpleSubject、開店年數=tpYear
          額外：offerWeight、quantitySumMonth、superNewProduct（新品標記）
v2 DOM 搜尋網址：https://m.1688.com/offer_search/-6D7033.html?keywords={quote(kw)}
```
**官方 API 不用**：GitHub `jiyun/1688` 只適用詳情頁非搜索頁；`open.1688.com` 的 `alibaba.search.product.search` 官方紅線禁止大規模全網商品採集，個人開發者權限極低 → 不採。

### 2.5 Output format / conventions

**Excel 欄位**
- v2（11 欄）：`序号`、`商品标题`、`主图(嵌图)`、`价格`、`价格min`、`价格max`、`销量`、`新上架(标x)`、`商品链接`、`品牌/标记`、`采集时间`。
- v3（16 欄）：在上述基礎上增加 `复购率%`、`重量kg`、`30天销量`、`新品標記`、`场景`（第 17 欄）等；第 2 欄加「品類」（合併多品類時）。
- **主圖 100% 嵌入**；C 欄嵌圖；I 列商品連結超連結；J 列標記「品牌／敏感詞／品牌+敏感詞」。
- **每 100 條切一表**（`表01`/`表02`…）。
- **列高**：標題列 32pt；**所有**資料列 120pt。
- **序號連續** 1..N。
- 圖片壓縮：`thumbnail(160,160)` LANCZOS + JPEG quality 72 + optimize=True（**不可用 `resize`**，會變形）。

**檔名規則**
- 批次檔：`<品類>-m1688-<YYYY-MM-DD>-bikeparts-<tag>-表NN.xlsx`
- v2 批次：`<品類>-m1688-<YYYY-MM-DD>-<tag>-表NN.xlsx`（例：`凯格尔-m1688-2026-08-28-v1-表01.xlsx`）
- 合併：`<品類>-m1688-<YYYY-MM-DD>-合併總表.xlsx`
- 最終：`<品類>-m1688-<YYYY-MM-DD>-最終總表.xlsx`
- 過濾鏈後綴：`-去衣褲` / `-去鞋` / `-去雜物` / `最終`
- **一律簡體中文檔名**（與 1688 來源語言一致；Windows PowerShell 顯示自動簡轉繁是假象）。
- 使用者手動加 `1` 前綴表示「看過」，**不要主動加**。

**已交付總表（實測數據）**

| 檔案 | 筆數 | 價格範圍 | 銷量範圍 | 圖片 | 備註 |
|---|---|---|---|---|---|
| `自行车配件-m1688-2026-10-08-最終總表.xlsx` | 1,141 | 10.0–199.5 | — | 1141/1141 | 刪 78 筆進口／知名大牌 |
| `自行车配件-m1688-2026-10-09-今日總表.xlsx` | 840 | 10–199 | 1–27,676 | 840/840 | 52MB→4.67MB |
| `健身辅助器具-m1688-2026-10-19-最終總表.xlsx` | 743→706 | 15–200 | 60–34,572 | 706/706 | 刪 37 筆非運動品項 |
| `儿童滑板车-m1688-2026-10-20-最終總表.xlsx` | 193 | 80–320 | 101–28,701 | 193/193 | 16 欄 |
| `户外便携休息用品-m1688-2026-10-20-最終總表.xlsx` | 262 | 80–330 | 51–18,459 | 262/262 | 16 欄 |
| `户外杯具茶具咖啡器具-m1688-2026-10-20-最終總表.xlsx` | 974 | — | — | 974/974 | 5,350,575 bytes |
| `保温饭盒汤罐保温箱-m1688-2026-10-20-最終總表.xlsx` | 267→270 | 20.0–220.0 | 200–94,054 | 270/270 | 16 欄 |
| `摩托车后尾箱-m1688-2026-10-20-最終總表.xlsx` | 156 | — | — | 156/156 | 已被使用者刪除 |
| `摩托车后靠背-m1688-2026-10-20-最終總表.xlsx` | 30 | — | — | 30/30 | 已被使用者刪除 |
| `鞋类-m1688-2026-10-20-最終總表.xlsx` | 566 | 10.0–185.2 | 101–155,579 | 566/566 | 舞蹈鞋 404／溯溪鞋 126／攀岩鞋 36 |
| `居家垃圾桶-m1688-2026-10-20-最終總表.xlsx` | 159 | 25.0–98.0 | 504–42,119 | 159/159 | 885,002 bytes |
| `居家店铺垃圾桶-m1688-2026-10-20-最終總表.xlsx` | 116 | 25.0–257.1 | 300–9,257 | 116/116 | 634,438 bytes；與 trash 159 筆重疊 0 |

**品類參數表（v3 CONFIG，明確記載）**

| cat | 品類 | 價格 | 銷量 | 品牌界線 | 特殊規則 |
|---|---|---|---|---|---|
| `gymaux` | 健身輔助器具 | 15–200 | ≥60 | 排除健身品牌 | 承重/重物剔除（單價 ≥80 例外保留）；服飾剔除但保留手套與襪子；敏感/醫療詞剔除；原使用者定價門檻**維持 15 元不變**（已拒絕降到 10 元的建議） |
| `kidscooter` | 兒童滑板車 | 80–350 | ≥100 | 只排國際大牌＋頭部品牌 | 電動滑板車排除；頭盔/鞋/手套不排除；非電動滑板車保留（regex `(?<![非无])电动`） |
| `outdoor` | 戶外便攜休息用品 | 80–350 | ≥50 | Coleman/Naturehike/Snow Peak/Mobi Garden/KingCamp/Blackdeer/Helinox/Therm-a-Rest/Sea to Summit/Intex/迪卡侬 + 台灣 CAMPFIRE/Chill Outdoor/閒風藍/TILLAK/SWS WELLHIKE/露營先生 | 帳篷/天幕/遮陽棚先剔除 62 筆 |
| `cup` | 戶外杯具/茶具/製茶製咖啡器具 | 20–150 | ≥100 | 國際＋國內大品牌都排除 | `cup_rules.py`：43 條 brand_rules、443 條 nonfit_rules、keep_rules=`六君子`+`焖烧|闷烧`；刀具放行 |
| `insul` | 保溫飯盒/湯罐/保溫箱 | 20–150（實際 20–220） | ≥100 | 同上 | 刻意保留奶茶桶/豆漿桶/冰粉桶/商用食堂保溫桶/郵政泡沫保溫箱/焖燒罐；刻意刪酒吧香檳冰桶/保冷軟袋/釣魚活魚桶/餌料盆/插電車載冰箱/木質單層飯盒/保溫杯水壺 |
| `mcase` | 機車後尾箱 | 40–240 | ≥100 | 車廠規則 | 25 筆哈雷改裝件使用者要留（**待實作**） |
| `mback` | 機車後靠背 | 30–180 | ≥100 | 車廠規則 | 含電動機車/電動車/電動三輪車 |
| `dance` | 舞蹈鞋 | 10–200 | ≥100 | 排除知名品牌 | `_DANCE_WRONG_SHOE` 只擋德训鞋/雨靴/馬丁靴/雪地靴/洞洞鞋/棉拖/穆勒鞋/樂福鞋 |
| `climb` | 攀岩鞋 | 10–200 | ≥100 | — | 兒童款保留；must_words 需同時含場景詞與品項字 `(?=.*鞋)` |
| `water` | 溯溪鞋 | 10–200 | ≥100 | — | 與 climb 重疊時**溯溪優先** |
| `trash` | 居家垃圾桶 | 25–105 | ≥500 | 不要知名品牌 | 風格須精緻/可愛/個性/涂鴉/卡通；`高颜值` 必列 |
| `trash2` | 居家店鋪垃圾桶 | 25–300 | ≥300 | 同上 | `dedup_against` 對 trash 去重（offerId + 標準化標題雙重比對） |

**品類產出規律（實測，選詞時必讀）**
- 高產出：護指手套 75/單詞（全場最佳：150 張卡→124 過價格→42 過銷量）、舉重腰帶 44、瑜伽磚軟木 44、瑜伽輪 32、普拉提圈 30、引體向上助力帶 30、大腿綁帶 23、臥推墊 20。
- 零產出（低價耗材）：彈力帶（137 張卡過價格 0 張）、拉力帶（117→2）、護膝（97→3）、瑜伽磚（112→18）、瑜伽伸展帶（87→2）、瑜伽帶（117→3）——1688 主流價 ￥3–14，低於 15 元門檻。
- 「半指手套」只 5 筆（搜尋下多是便宜款），與「護指手套」差異極大。
- 壞詞勿再跑：艾扬普拉提球（0 張，「艾扬」是大陸拼音廠商名）、握力辅助抓钩（8 張）、硬拉垫（23 張）、負重拉带（12 張）。
- 帶屬性字的關鍵詞幾乎無產出（攀岩鞋排扣/自由/全能/飛檐合計 0 筆）→ 少放此類詞。
- 護肩 64 筆被敏感規則擋掉是**正確的**（1688 混入肩周炎頸椎病復健器材，Shopee 禁賣風險）；已驗證健身護肩墊/舉重護肩/瑜伽護肩墊正常放行。

### 2.6 Verification & acceptance

**最終總表必驗六項（任一不通過不得交付）**
1. 圖片**逐列位元組相符**（N/N）
2. **序號連續** 1..N
3. 價格**零越界**（min ≤ x ≤ max）
4. 銷量**零越界**
5. 列高：標題 32pt / 資料列全部 120pt
6. **空連結 0**；品牌/IP 殘留掃描 0

**驗證方式（決定性）**
- 圖片 row→media 必須從 **xlsx 內部 drawing XML + rels 推導**；**不可用 openpyxl `Image.path`**（openpyxl 3.1.5 所有 `Image.path` 都指向 `/xl/media/image1.jpeg` → 圖片全錯）。
- 不可用 openpyxl 自我驗證（會互相抵消）。
- 若圖片重新壓縮，md5 比對必失敗 → 須用**同樣壓縮參數**再比對。
- 驗證腳本必須**獨立重新推導期望清單，不 import 建置邏輯**（`verify_shoes.py` 模式）。

**已驗證案例**：騎行服合併——來源 331 → URL 重複 0 → 標準化標題重複 9 → 最終 322 → 圖片 322/322 → 2,918,837 bytes。
**三輪 PDD/1688 對照驗證**：自行車配件 1,237/1,237 逐列正確；健身 743/743；cup 974；insul 267/267；總計六份已交付總表 3,543 筆連結皆無空值。

### 2.7 Known pitfalls & failure modes

| 類別 | 問題 | 處理 |
|---|---|---|
| 圖片 | openpyxl `Image.path` 全指向 image1.jpeg → 圖片與列錯配 | 改用 drawing XML + rels 推導 row→media |
| 圖片 | 合併時先刪提取目錄 → 圖片變 0 | 保留提取目錄後才 322/322 成功 |
| 圖片 | `resize(160,160)` 強制正方形 → 變形 | 改 `thumbnail(160,160)` |
| Excel | `ws.append([..., None, ...])` 丟失 inline string → **col3 標題全 None**（363 條全空） | 改 `ws.cell(r, c, value)` 逐 cell 寫 |
| Excel | `ws.append` 疑似對含 None/空字串的 list 丟部分 cell | 逐 cell 寫 |
| Excel | 只設定第 1–101 列 120pt，第 102–1238 列無設定 | 對**所有**資料列統一設 120pt |
| Excel | 讀 Excel 時遇到 `~$` 開頭鎖定檔 → Permission denied | 跳過 `~$` 檔 |
| 腳本 | **hardcode 漏改**：握力器 v2 寫進 v1 表覆蓋 | 用 `STATS_SUFFIX.replace('_','-')`；教訓寫入 memory（line 381） |
| 腳本 | 續跑 merge **重寫舊表** → 丟舊圖（fitness v1、釣魚椅 v7 受影響） | 續跑寫**新表**，不合併舊表 |
| 腳本 | 簡繁混雜 regex bug（註解用繁體「詞」U+8A5E、主程式 log 用簡體「词」U+8BCD） | 搜尋時要用對應編碼 |
| 腳本 | `SLEEP_BETWEEN_KW` 定義但未使用 → 每詞約 17 秒連打 | 補上使用 |
| 腳本 | 捲動迴圈從未真正捲動 | 修腳本 |
| 腳本 | `page.evaluate` 無逾時機制 → 滑塊讓採集永久卡死 | 看門狗 300s `os._exit(3)` ＋ `page.screenshot()` 診斷（渲染程序沒卡但 JS 執行緒被擋） |
| 腳本 | `(?i)` 不在 regex 最前方 → `re.PatternError: global flags not at the start` | 必須置於最前 |
| 規則 | 排除詞誤殺：「收纳桶」誤殺 83 筆（賣家寫「垃圾桶兼收納桶」） | 規則放寬重跑同一批詞，多撈回 58 筆 |
| 規則 | 規格詞 ≠ 品類特徵：分类桶/雙桶/大容量垃圾桶/收納桶/儲物桶/置物桶都是賣家描述 | 不擋 |
| 規則 | 風格詞誤當敏感詞：中古風/工業風是風格 | 改放寬 |
| 規則 | 「高颜值」漏掉 → 少一半好品（中文賣家用「简约/创意/高颜值」表達「精致」） | 補入 `_STYLE_WORDS` |
| 規則 | `_DANCE_WRONG_SHOE` 曾含 瑪麗珍/豆豆鞋/帆布鞋/運動鞋/健身鞋/瑜伽鞋 → **一次誤殺 88 筆正品** | 收斂為 8 類 |
| 規則 | 裸「收纳」誤殺手工芭蕾舞鞋 | 改具體「收纳包/袋/箱/盒」 |
| 規則 | 裸「鞋袜」被襪子規則殺 | `(?<!鞋)袜子` 負向斷言 |
| 規則 | 裸「車」導致汽車品項大量混入 | must_words 車輛語境改用精準載體詞 ∪ 車廠名 ∪ 車型名 |
| 規則 | 負向斷言未錨定 `^` → `re.search` 從每個位置嘗試 | 錨定或改精準詞 |
| 規則 | 相容性豁免把「改装/安装」算進相容文案 → 車廠規則形同虛設 | 移除 |
| 規則 | 「靠腰/护腰」在 1688 = 騎士防摔護具，不是車用靠背 | 視為最大詞彙陷阱 |
| 規則 | 「现货/現貨」是行銷詞不是品類特徵 | 從敏感詞移除 |
| 規則 | 「車載」不擋（意思是「車上用的」） | 保留 |
| 規則 | must_words 只要求場景詞不要求品項字 → 漏網 6 筆裝備 | 補 `(?=.*鞋)` |
| 規則 | 「小冰箱/車載冰箱」不能當電器排除（戶外保冷箱行銷說法） | 真正的電動保冷箱改用 12V/壓縮機/半導體判斷 |
| 品牌 | 品牌排除只寫繁體 → 漏抓 124 筆 | 同時寫簡體＋繁體 |
| 品牌 | 整車比對誤用「折叠车/通勤车/助力自行车」 | 零件標題常含相容性描述，不可刪 |
| 去重 | 跨任務池會**永久封鎖**被篩掉的品項 | v3 改 `dedup_mode:"delivered"` |
| 去重 | 重複商品其中一筆無圖片網址 | 優先保留有圖片網址的那筆 |
| 去重 | 採集器只擋 offerId → 同品類第二彈需在複審階段用標準化標題再擋一次 | `dedup_against`（finalize 與 verify 必須同步） |
| 風控 | 桌機版 tab 佔著會誘發風控 | 用 playwright 關掉桌機版 tab，開 m.1688.com 新 tab |
| 風控 | 腳本 abort 後 page[0] 留 `_____tmd_____` → 下次啟動風控檢查失敗 | 啟動前先確認 Chrome 狀態；等 30–60 分鐘自動解 |
| 風控 | 關閉所有 Chrome tab → Chrome 自己關閉 + `ECONNREFUSED` | 只關 `_____tmd_____` tab |
| 風控 | 首頁能開不代表搜尋頁恢復 | 恢復檢查抓 `a[href*="/offer/"]` 商品卡數量 |
| 風控 | 新 profile（9224）同帳號同樣被拒 | 帳號+IP 層級限制，換瀏覽器無效 |
| 風控 | 診斷腳本測配額會吃掉配額 | 不用 |
| 背景任務 | 任務變 `lost`（長時間閒置被回收），但 auto-save 批次檔仍在 | `merge_v3.py` glob 所有批次檔撈回（ts3 撿回 83 筆） |
| 背景任務 | 批次超時 | 每批 < 3300s，多切批 |
| 中斷 | 凱格爾 v3 腳本已寫但未啟動（缺 `Start-Process` + cron create） | 啟動要兩步都做 |
| 環境 | Chrome 9223 過 24 小時會死 | 啟動前檢查 |
| 環境 | 關閉指定實例會誤殺使用者個人 Chrome | 依 CommandLine 過濾 `remote-debugging-port=9223` |

### 2.8 Category-by-category pipeline history（v2 DOM 時代，已驗證）

| 日期 | 任務 | 條數 | 狀態 |
|---|---|---|---|
| 8-18 | 露營車 v1 / 釣魚椅 v1 | 33 / 52 | 完成 |
| 8-18 | 露營車 v2 / 釣魚椅 v8 | 23 / 19 | 完成 |
| 8-23 | 健身小器械 v1+續跑×2+fix-images | 139→140 | 表01 140 條／92 圖 |
| 8-24 | 運動手套 v1+續跑×2 | 245 | 表01–04，245 條／245 圖 |
| 8-25 | 護腕+頭巾 v1+續跑 | 102 | 表01–02 |
| 8-25 | MOLLE 戰術 v1 | 455 | 5 表，455 圖 100% |
| 8-25 | 帳篷 v1 | 201 | 完成 |
| 8-25 | 魚竿架 v1 | 142 | 完成 |
| 8-27 | 騎行褲 v1 | 280 | 完成（當時最高） |
| 8-27 | 兒童防護 v3 | 251 | 完成（v1/v2 中斷續跑） |
| 8-27 | 棒球棍 v1 | 100 | 完成（攻擊性過濾 96 條） |
| 8-27 | 防水包 v1 / v2 | 48 / 15 | 中斷（詞 27 / 39 觸風控） |
| 8-28 | 凱格爾 v1 / v2 / v3 | 41 / 93 / 49 | 合併總表 183 |
| 8-31 | 按摩舒緩 v1 / v2 | 373 / 307 | 680 → 清理 332 → 348 |
| 9-4 | 握力器 v1 / v2 | 160 / 102 | 合併 262 |
| 9-6 | 魚線 v1 | 162 | 清理 26 條非釣魚 → 136 |
| 9-15 | 特技棒球 v1/v1c/v1d/v2/v2b | 29/6/104/5/168 | 合併 312 → 過濾鏈 → 339 最終 |
| 9-18 | 臂力棒 v1 / v2 | 44 / 81 | 合併 125 |
| 9-18 | 刀具收藏 v1/v1b/v1d/v2/v2b | 127/85/46/118/68 | 合併 444（5.2MB） |
| 9-18 | 蝴蝶刀+爪子刀 v1 | 79 | v1b + v2 待跑 |
| 10-06 | 騎行服 v1/v1b/v1c | 233/97/1 | 331 → 去重 → 322 |

---
