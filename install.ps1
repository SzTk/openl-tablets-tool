# OpenL Tablets Tool インストールスクリプト (Windows PowerShell)
# 使い方: .\install.ps1

$ErrorActionPreference = "Stop"

$SkillName = "openl-edit"
$SkillDst  = "$env:USERPROFILE\.claude\skills\$SkillName"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== OpenL Tablets Tool インストール ===" -ForegroundColor Cyan

# 1. openl コマンドをグローバルインストール
Write-Host "[1/2] openl コマンドをインストール中..."
uv tool install $ScriptDir --force
# PATH に uv tools ディレクトリを追加（セッション内）
$uvBin = "$env:USERPROFILE\.local\bin"
if ($env:PATH -notlike "*$uvBin*") {
    $env:PATH = "$uvBin;$env:PATH"
}
# シェル起動時にも有効にする
uv tool update-shell 2>$null
Write-Host "      openl コマンド: インストール完了" -ForegroundColor Green

# 2. Claude Code スキルをインストール
Write-Host "[2/2] Claude Code スキルをインストール中..."
New-Item -ItemType Directory -Force -Path $SkillDst | Out-Null
Copy-Item "$ScriptDir\skills\$SkillName\SKILL.md" "$SkillDst\SKILL.md" -Force
Write-Host "      スキル配置先: $SkillDst\SKILL.md" -ForegroundColor Green

Write-Host ""
Write-Host "✅ インストール完了" -ForegroundColor Green
Write-Host ""
Write-Host "使い方:"
Write-Host "  openl read  <file.xlsx>     # Excel -> JSON"
Write-Host "  openl write <file.json>     # JSON -> Excel"
Write-Host "  openl roundtrip <file.xlsx> # 動作確認"
Write-Host ""
Write-Host "Claude Code から:"
Write-Host "  /openl-edit <file.xlsx>"
