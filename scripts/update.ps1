<#
.SYNOPSIS
  02324 离散数学 · 一键更新并发布思维导图
.DESCRIPTION
  1. 从 WPS 云文档下载最新笔记 (.docx)
  2. 自动生成 index.html
  3. 推送到 GitHub → Actions 自动部署到 Pages
#>

$SCRIPT_DIR = Split-Path -Parent $PSCommandPath
$ROOT = Split-Path -Parent $SCRIPT_DIR

Write-Host "╔══════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  02324 离散数学 · 思维导图更新工具  ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── 步骤选择 ──
Write-Host "请选择执行步骤:" -ForegroundColor White
Write-Host "  1) 全流程（下载 → 生成 → 推送）" -ForegroundColor Cyan
Write-Host "  2) 仅 生成 → 推送（跳过下载）" -ForegroundColor Cyan
Write-Host "  3) 仅 推送（跳过下载和生成）" -ForegroundColor Cyan
Write-Host "  0) 退出" -ForegroundColor Gray
$choice = Read-Host "请选择 (0-3)"

switch ($choice) {
    "0" { Write-Host "已退出" -ForegroundColor Gray; exit 0 }
    "2" { $skipDownload = $true; $skipGenerate = $false }
    "3" { $skipDownload = $true; $skipGenerate = $true }
    default { $skipDownload = $false; $skipGenerate = $false }
}

Write-Host ""

# ── Step 1: 下载 WPS 文档 ──
if (-not $skipDownload) {
    Write-Host "▶ Step 1/3: 从 WPS 下载最新笔记..." -ForegroundColor Yellow
    python "$SCRIPT_DIR\download_wps.py"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Step 1 失败，已退出" -ForegroundColor Red
        pause
        exit 1
    }
}

# ── Step 2: 生成 HTML ──
if (-not $skipGenerate) {
    Write-Host ""
    Write-Host "▶ Step 2/3: 生成思维导图 HTML..." -ForegroundColor Yellow
    python "$SCRIPT_DIR\generate.py"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Step 2 失败，已退出" -ForegroundColor Red
        pause
        exit 1
    }
}

# ── Step 3: 推送到 GitHub ──
Write-Host ""
Write-Host "▶ Step 3/3: 推送到 GitHub Pages..." -ForegroundColor Yellow

Push-Location $ROOT

git add index.html scripts/ update.bat .gitignore README.md
Write-Host "  git add ✓"

$commitMsg = "update: $(Get-Date -Format 'MM-dd HH:mm')"
git commit -m $commitMsg
if ($LASTEXITCODE -ne 0) {
    Write-Host "⚠️  commit 失败（可能没有变更内容）" -ForegroundColor Yellow
}

git push
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "❌ 推送失败（网络问题），重试请选 3) 仅推送" -ForegroundColor Red
    Pop-Location
    pause
    exit 1
}

Pop-Location

# ── 结果 ──
Write-Host ""
Write-Host "✅ 全部完成！" -ForegroundColor Green
Write-Host "GitHub Actions 正在自动部署到 Pages（约 1-2 分钟）" -ForegroundColor Cyan
Write-Host "访问: https://hanxiaodao.github.io/02324-mindmap/" -ForegroundColor White

pause
