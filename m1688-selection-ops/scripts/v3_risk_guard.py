# -*- coding: utf-8 -*-
"""v3_risk_guard.py — 1688 選品採集的「風控守門」(2026-11-04)

為什麼要這個模組 (2026-11-04 實測教訓):
  MiniMax Code 的選品流程本來是「單任務 15-30 詞, 任務之間隔數小時甚至隔天」,
  實戰日誌明載規律:
      「17-18 詞後開始限流, 18 詞時停 2-4 小時再跑, 或留到第二天」
      「每天 17 詞後限流, 第二天 m.1688 自然 reset」
      「後台跑 ≤ 20 詞 (15-25 分鐘)」
  但採集器本身沒有任何機制會阻止人把 100+ 詞一口氣跑完 ——
  2026-11-04 下午 48 詞/54 分鐘就觸發真風控, 全天 90 詞觸發 3 次。

四道閘門:
  1. 併發鎖   — 同一帳號同時只能有一個採集器 (你與我各跑各的會互相觸發風控)
  2. 輪次預算 — 每輪 ≤ ROUND_WORDS 詞, 到額自動停 (優雅停止 + 續跑記錄)
  3. 每日預算 — 每天 ≤ DAILY_WORDS 詞, 剩下的留到明天 (m.1688 隔天自然 reset)
  4. 產出熔斷 — 連續數輪產出量下滑 = 排序限流, 立刻停 (再跑只會出垃圾結果)

全部可用環境變數覆寫, 預設值對舊品類不影響:
  M1688_GUARD_ROUND_WORDS / M1688_GUARD_DAILY_WORDS / M1688_GUARD_ON
  M1688_GUARD_MIN_YIELD / M1688_GUARD_YIELD_WINDOW
"""
import json
import os
import shutil
import tempfile
import time
from datetime import date
from pathlib import Path

# 守門目錄可用 M1688_GUARD_DIR 覆寫 (單元測試必須用獨立目錄,
# 否則測試的清理會刪掉「正在使用中」的鎖與狀態檔)
GUARD_DIR = Path(os.environ.get(
    "M1688_GUARD_DIR", str(Path(tempfile.gettempdir()) / "1688_risk_guard")))
# 2026-11-10: 多通道 (手機帳號 / PC 帳號) 可平行 → 鎖與狀態檔要能分開
LOCK_SUFFIX = os.environ.get("M1688_LOCK_SUFFIX", "")
STATE_SUFFIX = os.environ.get("M1688_STATE_SUFFIX", "")
LOCK_DIR = GUARD_DIR / f"collect{LOCK_SUFFIX}.lock"


def _envi(name, default):
    try:
        return int(float(os.environ.get(name, default)))
    except Exception:
        return int(default)


def _envf(name, default):
    try:
        return float(os.environ.get(name, default))
    except Exception:
        return float(default)


def enabled():
    """守門總開關 (M1688_GUARD_ON=0 可完全關閉)"""
    return _envi("M1688_GUARD_ON", 1) != 0


def round_words():
    return _envi("M1688_GUARD_ROUND_WORDS", 20)      # 每輪詞數上限 (SOP: ≤15-20)


def daily_words():
    return _envi("M1688_GUARD_DAILY_WORDS", 60)      # 每日詞數上限 (SOP: 17 後須隔天)


def min_yield():
    return _envf("M1688_GUARD_MIN_YIELD", 3.0)       # 每詞平均產出低於此值 = 限流


def yield_window():
    return _envi("M1688_GUARD_YIELD_WINDOW", 6)      # 看最近幾輪的產出量


# ── 1. 併發鎖 ──────────────────────────────────────────────────────────
def _pid_alive(pid):
    """檢查行程是否存活 (絕對不能有任何「順手把它殺掉」的副作用)

    ⚠️ Windows 上 `os.kill(pid, 0)` 不是「檢查存活」——
       它的實作是 TerminateProcess(handle, 0) → 會**真的把行程殺掉**。
       2026-11-05 實測: 這裡若用 os.kill 做存活檢查, 等於用鎖的機制
       去殺掉正在跑採集器的行程 (2026-11-05 差點誤殺使用者的 g3 批次)。
       → Windows 一律用 OpenProcess + GetExitCodeProcess 查詢, 不送任何訊號。
    """
    try:
        pid = int(pid)
    except Exception:
        return False
    if pid <= 0:
        return False
    if os.name == "nt":
        try:
            import ctypes
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            k = ctypes.windll.kernel32
            h = k.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if not h:
                return False
            try:
                code = ctypes.c_ulong()
                if not k.GetExitCodeProcess(h, ctypes.byref(code)):
                    return False
                return code.value == STILL_ACTIVE
            finally:
                k.CloseHandle(h)
        except Exception:
            return False
    try:
        os.kill(pid, 0)
        return True
    except Exception:
        return False


def acquire_lock(kwset, tag="", pid=None):
    """取得採集鎖; 已被別人持有 → 回傳 False (呼叫端應中止, 不要硬敲)"""
    if not enabled():
        return True
    GUARD_DIR.mkdir(parents=True, exist_ok=True)
    pid = pid or os.getpid()
    info = {
        "kwset": kwset, "tag": tag, "pid": pid,
        "started": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    for attempt in range(2):
        try:
            LOCK_DIR.mkdir(parents=True, exist_ok=False)
            (LOCK_DIR / "owner.json").write_text(
                json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
            return True
        except FileExistsError:
            # 鎖存在 → 檢查持有者是否還活著 (用安全的 _pid_alive, 絕不殺行程)
            owner = {}
            try:
                owner = json.loads((LOCK_DIR / "owner.json").read_text(encoding="utf-8"))
            except Exception:
                owner = {}
            alive = _pid_alive(owner.get("pid"))
            if not alive and attempt == 0:
                # 殘留鎖 → 清掉重搶
                try:
                    shutil.rmtree(LOCK_DIR, ignore_errors=True)
                except Exception:
                    pass
                time.sleep(0.5)
                continue
            return False
    return False


def release_lock():
    try:
        shutil.rmtree(LOCK_DIR, ignore_errors=True)
    except Exception:
        pass


def lock_owner():
    try:
        return json.loads((LOCK_DIR / "owner.json").read_text(encoding="utf-8"))
    except Exception:
        return None


# ── 2/3. 輪次與每日預算 (狀態檔) ───────────────────────────────────────
def _state_file(kwset):
    return GUARD_DIR / f"session_{kwset}{STATE_SUFFIX}.json"


def _bydate_file(kwset):
    # 測試時用獨立前綴, 避免測試寫壞真正的續跑記錄
    pref = os.environ.get("M1688_BYDATE_PREFIX", "1688_done_v3_")
    return Path(tempfile.gettempdir()) / f"{pref}{kwset}_bydate.txt"


def count_words_today_from_bydate(kwset):
    """從「按日期」的續跑記錄算今天已採幾詞 (含守門模組上線前採的)

    2026-11-04: 守門是當天才加的, 若只看自己的 state 檔,
    當天先前已跑的 116 詞會被當成 0 → 預算形同虛設
    """
    p = _bydate_file(kwset)
    if not p.exists():
        return 0
    today = date.today().isoformat()
    n = 0
    try:
        for line in p.read_text(encoding="utf-8", errors="replace").splitlines():
            if line.startswith(today + "\t"):
                n += 1
    except Exception:
        return 0
    return n


def load_state(kwset):
    try:
        st = json.loads(_state_file(kwset).read_text(encoding="utf-8"))
    except Exception:
        st = {}
    if st.get("date") != date.today().isoformat():
        st = {"date": date.today().isoformat(), "words": 0, "requests": 0, "rounds": 0}
    # 若 state 檔還沒有今天的資料 → 用按日期記錄補上 (涵蓋守門上線前採的詞)
    if not st.get("seeded"):
        extra = count_words_today_from_bydate(kwset)
        if extra > st.get("words", 0):
            st["words"] = extra
        st["seeded"] = True
        st["seeded_from_bydate"] = extra
        save_state(kwset, st)
    return st


def save_state(kwset, st):
    try:
        GUARD_DIR.mkdir(parents=True, exist_ok=True)
        _state_file(kwset).write_text(
            json.dumps(st, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def budget_check(kwset, round_used):
    """回傳 (是否可繼續, 原因字串)"""
    if not enabled():
        return True, ""
    st = load_state(kwset)
    dcap = daily_words()
    rcap = round_words()
    if st.get("words", 0) >= dcap:
        return False, (f"已達每日預算 {dcap} 詞 (今日已採 {st['words']}); "
                       f"m.1688 隔天自然 reset, 剩下的留到明天")
    if round_used >= rcap:
        return False, (f"已達本輪預算 {rcap} 詞; "
                       f"依 SOP 17-18 詞後須停 2-4 小時或隔天, 請稍後續跑")
    return True, ""


def note_word(kwset, requests=0, got=0):
    st = load_state(kwset)
    st["words"] = st.get("words", 0) + 1
    st["requests"] = st.get("requests", 0) + int(requests or 0)
    st["got"] = st.get("got", 0) + int(got or 0)
    save_state(kwset, st)
    return st


def note_round(kwset):
    st = load_state(kwset)
    st["rounds"] = st.get("rounds", 0) + 1
    save_state(kwset, st)
    return st


def state_summary(kwset):
    st = load_state(kwset)
    return (f"今日 {st.get('words', 0)} 詞 / 預算 {daily_words()} 詞; "
            f"累計請求 {st.get('requests', 0)}; 輪次 {st.get('rounds', 0)}")


# ── 4. 產出熔斷 (排序限流偵測) ─────────────────────────────────────────
def yield_circuit(kwset, recent_yields, recent_raw=None):
    """recent_yields: 由舊到新的「每詞產出量(過濾後)」清單
    recent_raw:    對應的「每詞原始筆數」清單 (可選)

    必須區分兩種「產出變差」, 否則會誤殺長尾關鍵詞:
      (a) 帳號級排序限流: 原始仍有 ~180 筆, 但過不了價帶/銷量 → 真的該停
      (b) 關鍵詞本身稀疏: 原始就只剩 6–60 筆 (平台上沒這種商品) → 不該停,
          換個詞或接受低產出即可
    → 只有 (a) 同時成立時才熔斷: 平均產出低於門檻 **且** 平均原始筆數仍正常
    """
    if not enabled():
        return False, ""
    w = yield_window()
    if len(recent_yields) < w:
        return False, ""
    tail = recent_yields[-w:]
    thr = min_yield()
    avg_pass = sum(tail) / len(tail)
    if avg_pass >= thr or tail[-1] >= thr:
        return False, ""
    # 原始筆數也塌了 → 是關鍵詞稀疏, 不是帳號限流
    if recent_raw and len(recent_raw) >= w:
        raw_tail = recent_raw[-w:]
        avg_raw = sum(raw_tail) / len(raw_tail)
        if avg_raw < 60:
            return False, ""
    return True, (f"產出量熔斷: 最近 {w} 詞平均 {avg_pass:.1f} 筆/詞 "
                  f"(門檻 {thr}); m.1688 排序限流中, 再跑只出垃圾結果")
