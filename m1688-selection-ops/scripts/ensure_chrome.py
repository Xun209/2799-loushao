# -*- coding: utf-8 -*-
"""ensure_chrome.py — 確保 1688 的 CDP Chrome 活著 (9223 手機版 / 9226 桌機版)

Chrome 閒置或最後一個分頁被關時會自己結束 → 採集前先檢查並拉起來。
用法: python ensure_chrome.py [9223|9226]
"""
import io
import subprocess
import sys
import time
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
BASE = r"C:\Users\夏晨茗\AppData\Local\Google\Chrome"
MOBILE_UA = ("Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36")

PROFILES = {
    "9223": ("User Data_1688", MOBILE_UA, "https://m.1688.com/", "430,950"),
    "9226": ("User Data_1688_pc", None, "https://s.1688.com/", "1400,900"),
    "9222": ("User Data_PDD", None, "https://mobile.yangkeduo.com/", "430,950"),
}


def alive(port):
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=3).read()
        return True
    except Exception:
        return False


def main():
    port = sys.argv[1] if len(sys.argv) > 1 else "9223"
    profile, ua, url, size = PROFILES.get(port, PROFILES["9223"])
    if alive(port):
        print(f"[OK] port {port} 已在跑")
        return 0
    args = [CHROME, f"--remote-debugging-port={port}",
            f"--user-data-dir={BASE}\\{profile}"]
    if ua:
        args.append(f"--user-agent={ua}")
    args += ["--no-first-run", "--no-default-browser-check",
             "--disable-features=Translate,OptimizationGuideModelDownloading",
             f"--window-size={size}", url]
    subprocess.Popen(args)
    for i in range(20):
        time.sleep(1.5)
        if alive(port):
            print(f"[OK] port {port} 已啟動 (profile={profile}"
                  f"{', 手機 UA' if ua else ''})")
            return 0
    print(f"[FAIL] port {port} 啟動失敗")
    return 1


if __name__ == "__main__":
    sys.exit(main())
