# -*- coding: utf-8 -*-
"""v3_page_state.py — 頁面狀態分類 + 風控處置決策表

★ 來源: GitHub `superjack2050/1688-cli` (MIT) 的 `src/session/page-state.ts`
  與 `src/session/recovery.ts` 思路移植 (2026-11-09)

為什麼要移植:
  我們原本只有「真風控 / 不是風控」二分, 結果把三種狀態混在一起:
    1) risk_challenge (滑塊驗證)  → 必須停, 等人拖滑塊, 絕不自動重試
    2) rate_limited   (軟限流)    → 退避 60-90 秒後可重試 1 次
    3) site_changed   (版面改版)   → 停, 要更新解析器, 別硬幹
  → 混在一起的下場就是「明明只是軟限流卻整批停掉」或「明明是滑塊卻一直重試」。

判定依據刻意分層 (這是關鍵):
  - 網址層 (最可靠):  被導到 /punish 或帶 x5secdata
  - 文字層:           滑塊/驗證/異常 等字樣
  - 標題層:           「验证码拦截」
  ⚠ 注意: 歷史經驗包說「x5secdata 會誤判」—— 指的是**在回應文字/參數裡**看到就當風控;
    拿來比**頁面網址**是準的 (那是真的被導到驗證頁)。兩者不可混用。
"""
import re

# ── 判定用的正規式 (照抄 1688-cli, 已驗證於 2026-11-09 帳號正常時) ──
LOGIN_RE = re.compile(r"(?:login\.(?:1688|taobao)\.com|passport\.1688\.com)", re.I)
RISK_URL_RE = re.compile(r"(?:/punish\b|x5secdata)", re.I)
RISK_TEXT_RE = re.compile(
    r"(滑块|滑塊|拖动.*验证|拖動.*驗證|安全验证|安全驗證|验证码|驗證碼|验证一下|驗證一下"
    r"|检测到.*异常|檢測到.*異常|环境异常|環境異常|nc[_-]?captcha|请完成验证|請完成驗證)"
)
RATE_LIMIT_TEXT_RE = re.compile(
    r"(访问频繁|訪問頻繁|请求频繁|請求頻繁|操作频繁|操作頻繁|稍后再试|稍後再試"
    r"|访问受限|訪問受限|流量异常|流量異常|系统繁忙|系統繁忙|RGV587|被挤爆|被擠爆)"
)

NORMAL = "normal_1688_page"
NOT_LOGGED_IN = "not_logged_in"
RISK_CHALLENGE = "risk_challenge"
RATE_LIMITED = "rate_limited"
SITE_CHANGED = "site_changed"
UNKNOWN = "unknown"


def classify_page_state(url="", title="", text=""):
    """→ (kind, indicators[])"""
    hay = f"{title or ''}\n{text or ''}"
    ind = []

    if LOGIN_RE.search(url or ""):
        return NOT_LOGGED_IN, ["login-url"]
    if RISK_URL_RE.search(url or ""):
        return RISK_CHALLENGE, ["risk-url"]
    if RATE_LIMIT_TEXT_RE.search(hay):
        return RATE_LIMITED, ["rate-limit-text"]
    if RISK_TEXT_RE.search(hay):
        return RISK_CHALLENGE, ["risk-text"]
    if re.search(r"\.1688\.com", url or "", re.I):
        return NORMAL, ["1688-url"]
    return UNKNOWN, ind


def recovery_decision(kind, headed=False):
    """→ dict(action, retryable, max_retries, cooldown_sec, exit_code, code, message, hint)

    決策表移植自 1688-cli `recovery.ts` 的 `recoveryDecisionFor()`:
      ★ risk_challenge 的 max_retries = 0 → **絕不自動重試** (他們明講:
        "Do not add silent retry loops around risk-control failures")
      ★ rate_limited 允許 1 次重試 + 60-90 秒退避 (跟我們的「排序限流熔斷」不同層)
    """
    if kind == NOT_LOGGED_IN:
        return dict(action="pause_for_manual_login", retryable=False, max_retries=0,
                    cooldown_sec=0, exit_code=3, code="NOT_LOGGED_IN",
                    message="登入狀態失效 (被導到登入頁)",
                    hint="重新登入 1688 後再跑")
    if kind == RISK_CHALLENGE:
        return dict(action="pause_for_manual_challenge", retryable=False, max_retries=0,
                    cooldown_sec=0 if headed else 300,
                    exit_code=4, code="RISK_CONTROL",
                    message="1688 跳出滑塊驗證",
                    hint="在手機端視窗手動拖一次滑塊, 驗證後的 session 可沿用")
    if kind == RATE_LIMITED:
        return dict(action="backoff", retryable=True, max_retries=1,
                    cooldown_sec=60, exit_code=9, code="RATE_LIMITED",
                    message="1688 軟限流 (訪問頻繁/流量異常)",
                    hint="等 60-90 秒後降速重試 1 次; 仍失敗就停手")
    if kind == SITE_CHANGED:
        return dict(action="fail_with_artifacts", retryable=False, max_retries=0,
                    cooldown_sec=0, exit_code=1, code="SITE_CHANGED",
                    message="1688 頁面/接口結構可能已改版",
                    hint="抓真人流量比對 payload (capture_live_request.py) 後更新解析器")
    if kind == UNKNOWN:
        return dict(action="retry_once_then_fail", retryable=True, max_retries=1,
                    cooldown_sec=1, exit_code=1, code="UNKNOWN_RECOVERY_FAILURE",
                    message="未知的 1688 操作失敗")
    return dict(action="none", retryable=False, max_retries=0, cooldown_sec=0,
                exit_code=0, code="OK", message="正常")


def classify_from_body(body_text="", url=""):
    """直接對 mtop 回應文字/錯誤訊息分類 (不需頁面)

    ★★ 2026-11-09 實測修正 (順序很重要):
      1688 的「限流」回應會**同時帶著 punish 網址**:
        {"ret":["FAIL_SYS_USER_VALIDATE","RGV587_ERROR::SM::哎哟喂,被挤爆啦,请稍后重试"],
         "data":{"url":".../_____tmd_____/punish?x5secdata=…&action=captcha&pureCaptcha="}}
      → 若先比 punish 網址, 會被誤判成「滑塊挑戰」而直接停機;
        但它其實是**限流** → 正確處置是「退避 60-90 秒重試」, 不是等人拖滑塊。
      ★ 這正是使用者在 2026-11-06 遇到「滑動驗證一直失敗」的真實原因:
        那不是解不開的挑戰, 是需要「等」的限流懲罰頁 → 拖滑塊沒有用。
      → 回應層: **先判限流文字, 再判 punish 網址**。
      → 頁面層 (classify_page_state): 維持 1688-cli 的順序 (網址層優先), 因為
        真的被導到 /punish 頁就是挑戰。
    """
    if RATE_LIMIT_TEXT_RE.search(body_text or ""):
        return RATE_LIMITED, ["rate-limit-text"]
    if RISK_URL_RE.search(url or "") or RISK_URL_RE.search(body_text or ""):
        return RISK_CHALLENGE, ["risk-url"]
    if not body_text:
        return NORMAL, []
    if RISK_TEXT_RE.search(body_text):
        return RISK_CHALLENGE, ["risk-text"]
    return NORMAL, []
