---
name: m1688-selection-ops
description: 1688（m.1688.com）選品採集的實戰作業規則：現行 mtop API v3 採集、帳號與配額硬約束、風控四步流程、降級判定、批次節奏與 jitter 修正、合併去重口徑、最終總表驗收六項與踩坑表。當任務要在 1688 抓商品、跑關鍵詞批次、處理風控或合併 1688 產出的 Excel 時使用。
whenToUse: 使用者提到 1688、m.1688.com、選品採集、關鍵詞批次、tb966392038、9223 Chrome、風控、滑塊、_____tmd_____、mtop 或要合併 1688 產出的 Excel 時使用。
---

# 1688 選品採集作業規則（會話蒸餾）

> 完整 SOP 與腳本說明在 `references/w2-1688-sop.md`（技能名 `m1688-selection-sop`）。
> 本技能是**會話累積的增量與硬規則**；逐條完整規格見本技能的
> `references/w1-1688-selection.md`（29KB：硬約束、流程、欄位口徑、驗收六項、
> 完整踩坑表、逐品類歷史）。

## 一、帳號與環境（硬約束）

- **單帳號**：`tb966392038`（userid `2206668197833`）——**不能換號**。
- CDP 固定 `http://127.0.0.1:9223`，profile 固定 `User Data_1688`。
- **不重啟 Chrome 實例**（重啟＝新 TCP 連線＋新 session，觸發「異地／新設備」額外風控層）。
  Chrome 逾 24 小時會死，此時才重啟，且不得中途中斷使用者的實例。
- **不同時開兩個 Chrome 實例**（會崩潰且 session 被關聯）。
- **不用診斷腳本測配額**（診斷本身即搜索請求，會吃掉配額）。
- **不做指紋偽造**（stealth／Canvas 加噪）、**不用代理 IP 池**、**不破解滑塊**——
  僅採「換 session」與「降低請求量、貼合平台節奏」的合規手段。
- **只關桌機版分頁**（`www.1688.com`／`work.1688.com`），**不可關掉所有分頁**
  （會導致 Chrome 自關＋`ECONNREFUSED`）；不可誤關 `login.taobao.com` 分頁。
- 採集必須使用**專用分頁** `https://m.1688.com/?_mc=1`，完全不動使用者自己的分頁
  （腳本不得含任何 `.close()` 分頁程式碼）。

Chrome 啟動（路徑含空格必須用**單字串** ArgumentList）：

```powershell
$chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$argline = '--remote-debugging-port=9223 --user-data-dir="C:\Users\夏晨茗\AppData\Local\Google\Chrome\User Data_1688" --no-first-run --no-default-browser-check --disable-features=Translate,OptimizationGuideModelDownloading --window-size=430,950'
Start-Process -FilePath $chrome -ArgumentList $argline
```

只關 9223 實例（不動使用者一般 Chrome）：

```powershell
Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
  Where-Object { $_.CommandLine -like '*remote-debugging-port=9223*' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
```

## 二、風控四步流程（已確立）

1. **拖滑塊**（用 `trigger_punish.py` 把滑塊頁留在畫面上，使用者手動拖）
2. **等 40–50 分鐘**
3. **使用者手動搜 3–5 個詞暖機**（`python prompt_manual_search.py 9223`，每詞停留 3–5 秒）
4. **啟動採集**（等使用者回覆「好了」才開跑）

- 觸發訊號：`_____tmd_____`／滑塊驗證／`访问被拒绝`／`sec_verify`／
  `Access Denied`／`captcha`／`verify-slider`／`滑块验证`
  → **立即停止，不自動跳頁重試、不自動破解**，交使用者手動處理。
- 拖完滑塊後**搜索配額僅 2–3 次**；**滾動（同頁 XHR）不消耗搜索配額**。
- 風控是**帳號 + IP 層級**：新開 profile（9224／`User Data_1688_b`）同帳號同樣被拒；
  隔夜 16 小時不碰 1688 也未恢復 → **換瀏覽器無效**。
- 腳本 abort 後 `page[0` 會留 `_____tmd_____`，導致下次啟動的風控檢查失敗 →
  啟動前先確認 Chrome 狀態。

## 三、降級（帳號級限流）判定

- **只有「連續 2 個詞 API 呼叫全失敗」才算降級**；
  不把「0 張卡」直接當壞詞（要先測多個詞）。
- 降級判定以**捲動後商品卡數**為準：首屏 11–12 張 = 降級；
  捲 5 次到 90+ 才算恢復（部分腳本用 `>40` 張，`MIN_CARDS_ENOUGH` 由 60 改 100）。
  **不能只看首屏。**
- 加 **6 秒複讀確認**避免瞬間性誤判。
- **連續 3 詞 <30 張卡 → 自動停止。**

## 四、採集請求形狀（現行 v3，已實測可用）

```
端點：h5api.m.1688.com/h5/mtop.1688.wap.seo.offer.get/1.0/
簽名：md5(_m_h5_tk前段 & t & 12574478 & data)
必須在頁面 context 內 fetch（沿用 Chromium TLS 指紋 + cookie，不觸發 document 層配額）
參數：sortType="price_asc"（pop 綜合排序命中率最差：4/60）、priceStart/priceEnd、pageSize=60
欄位口徑：銷量=bookedCount（非 saleQuantity）、復購率=shopRepurchaseRate（非 offerRepurchaseRate）、
          價格=price、標題=simpleSubject、開店年數=tpYear
          額外：offerWeight、quantitySumMonth、superNewProduct（新品標記）
```

- **舊的 DOM 捲動採集（`collect_m1688_v2_*.py`）已停止使用**（檔案保留）。
- **官方 API 不採**：`open.1688.com` 的 `alibaba.search.product.search` 官方紅線禁止
  大規模全網商品採集，個人開發者權限極低。

核心腳本（`E:\1shopee\1688選品工作流\scripts\`）：

```
collect_m1688_v3_api.py   採集（簽名／分頁／排序／價格帶／重試／逾時／續採詞表／去重池寫入）
v3_config.py → v3_rules.py  品類設定與共用規則引擎
merge_v3.py → finalize_v3.py → verify_v3.py → final_check_cat.py   三段式管線
trigger_punish.py / prompt_manual_search.py / check_env.py / check_punish.py
```

## 五、執行流程（v3）

1. 檢查 9223 Chrome 是否活著 → 2. 確認登入（`open_1688_only.py` 印 `lid` cookie）
→ 3. 檢查風控狀態（`check_env.py`；滑塊頁用 `trigger_punish.py` 保留）
→ 4. **暖機**（`prompt_manual_search.py 9223`，等使用者回「好了」）
→ 5. 詞表 `keywords_<cat>.txt` + `v3_config.py` 的 `CATS["<cat>"]`
→ 6. **雙向測試篩選規則（必跑）**：`test_rules.py` 確認應保留案例 0 誤擋、應刪除案例 0 漏擋
→ 7. 採集 `python collect_m1688_v3_api.py <start> <end> <tag> <kwset>`
（`<tag>` **必須是新值**，避免覆蓋既有批次檔）
→ 8. `merge_v3.py` → 9. `finalize_v3.py` → 10. `verify_v3.py`
→ 11. `final_check_cat.py` + `check_all_delivered.py`
→ 12. `audit_cat_rules.py <cat> full` 稽核 + `test_rules.py` 乾跑抽樣
→ 13. 交付（media 標籤 + 路徑 + 驗證數據 + 刪除明細 + **邊界保留項列表**）

## 六、選品條件（每任務的硬約束）

- **售價區間是硬約束**。已出現區間：6-21／6-36／6-46／8-48／9-49／9-99／10-30／
  10-40／10-80／10-100／10-150／10-200／10-300／20-150／25-105／25-300／30-200／
  40-240／50-300／80-350。
- **銷量門檻是硬約束**。MIN_SALES：1（凱格爾）／50（戶外）／60（健身）／100（多數）／
  300（trash2）／500（trash）。
- **品牌策略**：可接受中國中階牌；排除國際／進口大牌與國內頭部品牌。
  **品牌排除必須同時覆蓋簡體與繁體寫法**（「禧马诺／禧瑪諾」、「蓝图／藍圖」）——
  先前只寫繁體導致**漏抓 124 筆**。
- **禁賣品全域護欄**（`v3_rules.py` 的 `BANNED`，必須全域而非各品類各自寫）：
  护目镜／護目鏡／护目／護目／防护眼镜／防護眼鏡／保护眼镜／保護眼鏡／运动眼镜／
  運動眼鏡／眼镜／眼鏡／泳镜／泳鏡／滑雪镜／滑雪鏡／潜水镜／潛水鏡／风镜／風鏡／
  防风镜／防風鏡 ＋ `goggles` 小寫比對。
- **刀具不是禁賣品**（使用者 2026-10-20 定案），只擴充「製茶刀具」。
- 車廠規則（`brand_ctx_rules` + `compat_allow`）：車廠名是主詞就刪、寫「適用 XX」就留；
  **「改装／安装」不算相容字樣**。
- **只靠負向排除會一直漏網** → 每品類配 `must_words` 正面下限，且正面清單要夠寬。

## 七、節奏（隨機化，不可單一）

- 最保守版：詞間停頓 80–240s（拆 3–6 段不等長、±45% 浮動）；冷卻 150–380s／每 4–8 詞；
  長休 200–520s／40% 機率；每 2 詞回首頁；捲動 5–10 次；間隔 3.0–7.5s；
  `HARD_TIMEOUT_SEC = 220*60`。
- **jitter 修正（重要）**：拆段 sleep 全部累加會放大 3–6 倍（80–240s 變 525s）；
  須**權重歸一化**使總時長精確等於 base。修正後效率提升約 5 倍，
  風控耐受從 3–4 詞升到 **12–18 詞**。
- 100 詞/任務、100 分鐘硬切；auto-partial 每 20 詞留底；每 8 詞 cool-down 60s。

## 八、任務管理

- detach 執行（`Start-Process`），ETA 只設 **1 個 cron（once）**，不循環。
- 背景 shell 任務上限 3600 秒 → 採集**每批 < 3300 秒**，寧可多切幾批。
- 檔名 v2/v3 為續跑編號，**不修改舊腳本**（複製模板改 hardcode 派生新腳本）。
- 去重池：`E:\1shopee\excel\_existing_offerids_2026-08-27.txt` 單一檔持續累積。
  v3 新增 `dedup_mode:"delivered"`：只排除已交付總表 ∪ 本品類既有批次，
  避免舊任務被篩掉的品項**永久封鎖**。

## 九、去重與交付口徑

```
來源 N 條 → URL 重複 X → 標準化標題重複 Y → 最終 Z 條 → 圖片 Z/Z
```

- 已驗證案例（騎行服）：331 → 0 → 9 → **322** → 圖片 322/322 → 2,918,837 bytes。
- 重複商品**優先保留有圖片網址的那筆**。
- `dedup_against` 在 finalize 與 verify 必須同步，否則 verify 誤報。

## 十、最終總表驗收六項（任一不通過不得交付）

1. 圖片**逐列位元組相符**（N/N）
2. **序號連續** 1..N
3. 價格**零越界**（min ≤ x ≤ max）
4. 銷量**零越界**
5. 列高：標題 32pt / 資料列**全部** 120pt
6. **空連結 0**；品牌/IP 殘留掃描 0

## 十一、最容易踩的坑（精選）

| 問題 | 處理 |
|---|---|
| openpyxl `Image.path` 全指向 image1.jpeg → 圖片與列錯配 | 改用 **drawing XML + rels** 推導 row→media |
| 合併時先刪提取目錄 → 圖片變 0 | 保留提取目錄到存檔完成 |
| `ws.append([..., None, ...])` 丟失 inline string → col3 標題全 None | 改 `ws.cell(r, c, value)` 逐 cell 寫 |
| 只設第 1–101 列 120pt，第 102–1238 列沒設 | 對**所有**資料列統一設 |
| 續跑 merge 重寫舊表 → 丟舊圖 | 續跑寫**新表**，不合併舊表 |
| 排除詞誤殺（「收纳桶」誤殺 83 筆） | 放寬規則並**重跑同一批關鍵詞**撈回 |
| 裸「車」導致汽車品項混入 | must_words 車輛語境改用精準載體詞 |
| 「靠腰／护腰」在 1688 = 騎士防摔護具 | 最大詞彙陷阱，必須特判 |
| 「现货／現貨」是行銷詞不是品類特徵 | 從敏感詞移除 |
| regex `(?i)` 不在最前 | 會拋 `global flags not at the start`，必須置於最前 |
| `page.evaluate` 無逾時 → 滑塊讓採集永久卡死 | 看門狗 300s `os._exit(3)` |
| `Select-Object -First N` 截斷管線 → exit code 4294967295 | **非真錯誤**，別誤判 |
| 任務變 `lost` 但 auto-save 批次檔仍在 | `merge_v3.py` glob 所有批次檔撈回 |

## 十一之二、PC 通道（s.1688.com）專屬教訓 — 2026-11-25 實測

| 坑 | 症狀 | 修法 |
|---|---|---|
| ★ **`afterPrice` 是「運費」不是「銷量」** | PC 回應 `afterPrice = {"matKey":"freight_desc","text":"运费3.5元"}` → 正則抓到「3」→ **銷量被讀成 3** → 好商品被當「銷量不足」誤殺（實測 40 筆/5 詞） | 銷量改以 **`bookedCount` 為主**（與手機版口徑一致）；`afterPrice` 只有 `matKey` 非 `freight_desc`/`free_shipping` 時才採用 |
| 銷量欄位顯示「包邮」/「运费3.5元」 | Excel 銷量欄是人話但沒意義 | 運費類文字 → 改寫成 `成交{bookedCount}笔` |
| 復購率拿不到 | `afterTags.matKey=return_rate`（不是 title=回头率） | 退回用 `offerRepurchaseRate` |
| **PC 搜尋頁關鍵字必須 GBK 編碼** | UTF-8 → 亂碼、found=0~8、標題不相干 | `keywords.encode('gbk')` 後 percent-encode |
| PC 頁每詞只回 1 頁 | `[分頁] 沒有下一頁` → 每詞 120 筆候選 | 正常；靠多詞覆蓋 |
| PC 通道也要有守門 | 原本只鎖不預算 | 已接 `budget_check` / `note_word`（`_pc` 狀態檔） |

**選品經驗（居家桌板類）**：
- **「书桌／办公桌／写字台／课桌」在標題裡是「用途修飾語」不是品類主詞**
  → 放進排除清單會誤殺本類（實測「可折叠床上升降小桌子家用学习书桌卧室」被殺，單此一詞誤殺 114 筆）
  → **不要放進 `exclude`**，改由 `must_words` 當品類下限把關（不含本類品項字 → 自然被擋）
- **「用途修飾語」與「品類主詞」要用負向斷言區分**
  → 例「笔记本电脑桌板」是本類主推品，不能被「电脑桌」殺 →
  `(?<!笔记本)(?<!筆電)电脑桌(?!板)`
- **`keep_rules` 是「純 pattern 列表」，`deny_rules` 才是 `(pat, name)` tuple 組**（格式寫錯會拋 TypeError）
  → `keep_rules` 適合「本類品項字 + 居家語境 → 整條放行」，再用 `deny_rules` 否決戶外/商用
- 產出偏低時**先看剔除側檔的原因分布**（`%TEMP%\v3_<cat>_<tag>_dropped.txt`），
  不要直接放寬條件 —— 有時是欄位映射 bug 而不是規則太嚴

## 十二、已作廢的規則（別再用）

- DOM 捲動採集（`collect_m1688_v2_*.py`）→ 已被 mtop API v3 取代。
- `sortType="pop"` → 作廢，用 `price_asc`。
- `pageSize=200` → 作廢（會被砍回 60）。
- `saleQuantity` / `offerRepurchaseRate` → 作廢，用 `bookedCount` / `shopRepurchaseRate`。
- 跨任務去重池阻擋全部歷史 offerId → 作廢，改 `dedup_mode:"delivered"`。
- `open.1688.com` 官方 API → 不採。
