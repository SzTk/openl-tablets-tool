# OpenL Tablets Tool インストールスクリプト (Windows PowerShell)
# 使い方: .\install.ps1

$ErrorActionPreference = "Stop"

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$SkillsSrc  = "$ScriptDir\skills"
$SkillsDst  = "$env:USERPROFILE\.claude\skills"

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

function Install-Skill {
    param([string]$Name)
    $dst = "$SkillsDst\$Name"
    New-Item -ItemType Directory -Force -Path $dst | Out-Null
    Copy-Item "$SkillsSrc\$Name\SKILL.md" "$dst\SKILL.md" -Force
    Write-Host "      $Name -> $dst\SKILL.md" -ForegroundColor Green
}

Install-Skill "openl-tablets-edit"
Install-Skill "openl-tablets-create"
Install-Skill "openl-tablets-deploy"

# openl-lib（共有スキーマ）をコピー
$libDst = "$SkillsDst\openl-lib"
New-Item -ItemType Directory -Force -Path $libDst | Out-Null
Copy-Item "$SkillsSrc\openl-lib\SCHEMA.md" "$libDst\SCHEMA.md" -Force
Write-Host "      openl-lib -> $libDst\SCHEMA.md" -ForegroundColor Green

Write-Host ""
Write-Host "✅ インストール完了" -ForegroundColor Green
Write-Host ""
Write-Host "使い方:"
Write-Host "  openl read  <file.xlsx>     # Excel -> JSON"
Write-Host "  openl write <file.json>     # JSON -> Excel"
Write-Host "  openl roundtrip <file.xlsx> # 動作確認"
Write-Host ""
Write-Host "Claude Code から:"
Write-Host "  /openl-tablets-edit <file.xlsx>      # 既存 Excel を編集"
Write-Host "  /openl-tablets-create [output.json] # 新規作成"
Write-Host "  /openl-tablets-deploy <file.xlsx>    # Azure にデプロイ"
