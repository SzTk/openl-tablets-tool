#!/usr/bin/env bash
# OpenL Tablets Tool インストールスクリプト
# 使い方: bash install.sh

set -e

SKILL_NAME="openl-edit"
SKILL_DST="$HOME/.claude/skills/$SKILL_NAME"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=== OpenL Tablets Tool インストール ==="

# 1. openl コマンドをグローバルインストール
echo "[1/2] openl コマンドをインストール中..."
uv tool install "$SCRIPT_DIR" --force
uv tool update-shell 2>/dev/null || true
# PATH に uv tools ディレクトリを追加（現在のセッション内）
export PATH="$HOME/.local/bin:$PATH"
echo "      openl コマンド: OK ($(openl --version 2>/dev/null || echo installed))"

# 2. Claude Code スキルをインストール
echo "[2/2] Claude Code スキルをインストール中..."
mkdir -p "$SKILL_DST"
cp "$SCRIPT_DIR/skills/$SKILL_NAME/SKILL.md" "$SKILL_DST/SKILL.md"
echo "      スキル配置先: $SKILL_DST/SKILL.md"

echo ""
echo "✅ インストール完了"
echo ""
echo "使い方:"
echo "  openl read  <file.xlsx>          # Excel → JSON"
echo "  openl write <file.json>          # JSON → Excel"
echo "  openl roundtrip <file.xlsx>      # 動作確認"
echo ""
echo "Claude Code から:"
echo "  /openl-edit <file.xlsx>"
