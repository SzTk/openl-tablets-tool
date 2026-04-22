#!/usr/bin/env bash
# OpenL Tablets Tool インストールスクリプト
# 使い方: bash install.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_SRC="$SCRIPT_DIR/skills"
SKILLS_DST="$HOME/.claude/skills"

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

install_skill() {
  local name="$1"
  local dst="$SKILLS_DST/$name"
  mkdir -p "$dst"
  cp "$SKILLS_SRC/$name/SKILL.md" "$dst/SKILL.md"
  echo "      $name → $dst/SKILL.md"
}

install_skill "openl-edit"
install_skill "openl-new"

# openl-lib（共有スキーマ）をコピー
mkdir -p "$SKILLS_DST/openl-lib"
cp "$SKILLS_SRC/openl-lib/SCHEMA.md" "$SKILLS_DST/openl-lib/SCHEMA.md"
echo "      openl-lib → $SKILLS_DST/openl-lib/SCHEMA.md"

echo ""
echo "✅ インストール完了"
echo ""
echo "使い方:"
echo "  openl read  <file.xlsx>          # Excel → JSON"
echo "  openl write <file.json>          # JSON → Excel"
echo "  openl roundtrip <file.xlsx>      # 動作確認"
echo ""
echo "Claude Code から:"
echo "  /openl-edit <file.xlsx>          # 既存 Excel を編集"
echo "  /openl-new  [output.json]        # 新規作成"
