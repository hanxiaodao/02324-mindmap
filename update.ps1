<#
.SYNOPSIS
  02324 离散数学 · 一键更新并发布思维导图
.DESCRIPTION
  1. 从 WPS 云文档下载最新笔记 (.docx)
  2. 自动生成 index.html
  3. 推送到 GitHub → Actions 自动部署到 Pages
#>

$ROOT = Split-Path -Parent $PSCommandPath
$HAS_ERROR = $false

Write-Host "╔══════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  02324 离散数学 · 思维导图更新工具  ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: 下载 WPS 文档 ──
Write-Host "▶ Step 1/3: 从 WPS 下载最新笔记..." -ForegroundColor Yellow
python "$ROOT\download_wps.py"
if ($LASTEXITCODE -ne 0) {
    $HAS_ERROR = $true
}

# ── Step 2: 生成 HTML ──
Write-Host ""
Write-Host "▶ Step 2/3: 生成思维导图 HTML..." -ForegroundColor Yellow
python "$ROOT\generate.py"
if ($LASTEXITCODE -ne 0) {
    $HAS_ERROR = $true
}

# ── Step 3: 推送到 GitHub ──
Write-Host ""
Write-Host "▶ Step 3/3: 推送到 GitHub Pages..." -ForegroundColor Yellow

Push-Location $ROOT

git add index.html generate.py download_wps.py update.ps1 .gitignore README.md
Write-Host "  git add ✓"

$commitMsg = "update: $(Get-Date -Format 'MM-dd HH:mm')"
git commit -m $commitMsg
if ($LASTEXITCODE -eq 0) {
    Write-Host "  git commit ✓"
} else {
    Write-Host "  ⚠️  commit 失败（可能没有变更内容）" -ForegroundColor Yellow
}

git push
if ($LASTEXITCODE -eq 0) {
    Write-Host "  git push ✓"
} else {
    Write-Host "  ⚠️  git push 失败，检查网络连接后重试" -ForegroundColor Yellow
}

Pop-Location

# ── 结果 ──
Write-Host ""
if ($HAS_ERROR) {
    Write-Host "⚠️  流程走完，但有步骤出错，请查看上方日志" -ForegroundColor Yellow
} else {
    Write-Host "✅ 全部完成！" -ForegroundColor Green
    Write-Host "GitHub Actions 正在自动部署到 Pages（约 1-2 分钟）" -ForegroundColor Cyan
    Write-Host "访问: https://hanxiaodao.github.io/02324-mindmap/" -ForegroundColor White
}

pause
