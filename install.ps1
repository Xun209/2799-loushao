# install.ps1 — 把本技能裝進你的 client 的 skills 目錄
#
# 用法:
#   powershell -ExecutionPolicy Bypass -File .\install.ps1
#   powershell -ExecutionPolicy Bypass -File .\install.ps1 -Target "C:\path\to\skills" -Profile desktop
#
# 不帶參數時會自動猜測:
#   1) %USERPROFILE%\.dsh\profiles\<profile>\node_modules\dsh-minimax-workflows\skills
#   2) %USERPROFILE%\.minimax\workspace\skills
param(
    [string]$Target = "",
    [string]$Profile = "",
    [switch]$Check
)

$ErrorActionPreference = "Continue"
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
$skillDir = Join-Path $src "m1688-selection-ops"
# fallback: 若使用者只下載了技能資料夾 (本檔與技能同層)
if (-not (Test-Path $skillDir)) { $skillDir = $src }

Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  m1688-selection-ops 安裝" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  來源: $skillDir"

if ($Check) {
    Write-Host "`n[檢查模式] 不做任何寫入" -ForegroundColor Yellow
    if (Test-Path (Join-Path $skillDir "SKILL.md")) { Write-Host "  [OK] SKILL.md 存在" }
    if (Test-Path (Join-Path $skillDir "references\w1-1688-selection.md")) { Write-Host "  [OK] references 存在" }
    $n = (Get-ChildItem (Join-Path $skillDir "scripts") -File -ErrorAction SilentlyContinue | Measure-Object).Count
    Write-Host "  [OK] scripts 腳本 $n 支"
    Write-Host "`n  Python: $(& python --version 2>&1)"
    exit 0
}

# ── 猜目標目錄 ──
if (-not $Target) {
    $cands = @()
    $dshRoot = Join-Path $env:USERPROFILE ".dsh\profiles"
    if (Test-Path $dshRoot) {
        $profs = Get-ChildItem $dshRoot -Directory -ErrorAction SilentlyContinue
        if ($Profile) { $profs = $profs | Where-Object { $_.Name -eq $Profile } }
        foreach ($p in $profs) {
            $cands += (Join-Path $p.FullName "node_modules\dsh-minimax-workflows\skills")
        }
    }
    $cands += (Join-Path $env:USERPROFILE ".minimax\workspace\skills")
    $cands += (Join-Path $env:USERPROFILE ".dsh\skills")
    $Target = $cands | Where-Object { $_ -and (Test-Path (Split-Path $_ -Parent)) } | Select-Object -First 1
    if (-not $Target) { $Target = $cands[0] }
}

Write-Host "  目標: $Target"
if (-not (Test-Path $Target)) {
    Write-Host "  [建立目錄] $Target" -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $Target | Out-Null
}

$dst = Join-Path $Target "m1688-selection-ops"
Write-Host "`n[安裝] 複製到 $dst ..."
if (Test-Path $dst) {
    Write-Host "  目標已存在 → 先移除舊版" -ForegroundColor Yellow
    Remove-Item $dst -Recurse -Force -ErrorAction SilentlyContinue
}
Copy-Item $skillDir $dst -Recurse -Force
Write-Host "  [OK] 已安裝" -ForegroundColor Green

Write-Host "`n[驗證]"
$ok = $true
foreach ($f in @("SKILL.md", "references\w1-1688-selection.md", "README.md")) {
    if (Test-Path (Join-Path $dst $f)) { Write-Host "  [OK] $f" } else { Write-Host "  [缺] $f" -ForegroundColor Red; $ok = $false }
}
$sc = Get-ChildItem (Join-Path $dst "scripts") -File -ErrorAction SilentlyContinue
Write-Host "  [OK] scripts $((($sc | Measure-Object).Count)) 支"

Write-Host "`n[下一步]" -ForegroundColor Cyan
Write-Host "  1. 開 Chrome debug 實例（手機通道 9223 / PC 通道 9226）"
Write-Host "     powershell -ExecutionPolicy Bypass -File .\scripts\ensure_chrome.py 9223"
Write-Host "  2. 在該視窗手動登入 1688（帳號信譽＝風控分數，不可換號）"
Write-Host "  3. 環境自檢: python -u scripts\check_env.py"
Write-Host "  4. 開始採集: python -u scripts\collect_m1688_v3_api.py 1 20 a1 keywords_xxx.txt"
Write-Host ""
Write-Host "  完整說明: $dst\README.md"
Write-Host "====================================================" -ForegroundColor Cyan
