# -*- coding: utf-8 -*-
"""v3_artifacts.py — 採集異常時留證據 (截圖 + HTML + 網路摘要)

★ 來源: GitHub `superjack2050/1688-cli` 的 `src/session/artifacts.ts` /
  `src/session/recovery.ts` (attachTrace + captureFailureArtifact) 思路移植 (2026-11-09)

為什麼要:
  我們以前遇到「看門狗懸空 302 秒」「訪問被拒絕」只有一行 log,
  事後無法判斷到底是 風控 / 改版 / 網路 / 分頁被搶 → 只能靠猜。
  留 artifact 後, 下次看圖 + HTML 就能立刻分辨 (對應 1688-cli 的 site_changed 判定)。
"""
import io
import json
import sys
import time
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)

ART_DIR = Path(r"E:\1shopee\1688選品工作流\日誌\artifacts")


def capture(ctx, tag="", kind="", note="", max_net=40):
    """把當下頁面狀態 + 最近網路流量存檔。回傳 artifact 目錄。"""
    ART_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    d = ART_DIR / f"{stamp}_{tag or 'run'}{('_' + kind) if kind else ''}"
    d.mkdir(parents=True, exist_ok=True)
    info = {"stamp": stamp, "tag": tag, "kind": kind, "note": note}

    pages = [p for p in ctx.pages if not p.is_closed()]
    info["page_count"] = len(pages)
    if pages:
        pg = pages[-1]
        try:
            info["url"] = pg.url
            info["title"] = pg.title()
        except Exception as e:
            info["url_error"] = str(e)[:120]
        try:
            (d / "page.html").write_text(
                pg.evaluate("() => document.documentElement.outerHTML") or "",
                encoding="utf-8", errors="replace")
        except Exception:
            pass
        try:
            # 背景分頁用 Playwright screenshot 會失敗 → 改用 CDP Page.captureScreenshot
            cdp = ctx.new_cdp_session(pg)
            cdp.send("Page.enable")
            shot = cdp.send("Page.captureScreenshot", {"format": "png"})
            import base64
            (d / "shot.png").write_bytes(base64.b64decode(shot["data"]))
        except Exception as e:
            try:
                (d / "shot.png").write_bytes(pg.screenshot(timeout=8000))
            except Exception:
                pass
        try:
            (d / "text.txt").write_text(
                pg.evaluate("() => document.body ? document.body.innerText.slice(0,4000) : ''") or "",
                encoding="utf-8", errors="replace")
        except Exception:
            pass

    # 最近網路流量 (用 CDP)
    try:
        cdp = ctx.new_cdp_session(pages[-1]) if pages else None
        if cdp:
            net = []
            cdp.on("Network.responseReceived", lambda e: net.append({
                "status": (e.get("response") or {}).get("status"),
                "url": (e.get("response") or {}).get("url", "")[:220],
            }))
            cdp.on("Network.loadingFailed", lambda e: net.append({
                "failed": e.get("errorText"),
                "url": (e.get("requestId") or "")[:40],
            }))
            time.sleep(2.0)
            (d / "network.json").write_text(
                json.dumps(net[-max_net:], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

    (d / "info.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    return d


def main():
    """手動抓一次當前狀態 (除錯用)"""
    from playwright.sync_api import sync_playwright
    sys.path.insert(0, r"E:\1shopee\1688選品工作流\scripts")
    with sync_playwright() as p:
        b = p.chromium.connect_over_cdp("http://127.0.0.1:9223")
        ctx = b.contexts[0]
        d = capture(ctx, tag="manual", kind="manual", note="手動觸發")
        print(f"artifact -> {d}")
        for f in sorted(d.iterdir()):
            print(f"   {f.name:<16} {f.stat().st_size:>9,} bytes")
        b.close()


if __name__ == "__main__":
    main()
