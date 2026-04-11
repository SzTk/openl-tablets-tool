# OpenL Tablets Tool

OpenL Tablets 形式の Excel ファイルを JSON/YAML に変換・編集・書き戻しする CLI ツールと Claude Code スキルのセットです。

## 対応テーブル種別

| 種別 | 説明 |
|---|---|
| SimpleDecisionTable | 条件 → 結果のルールテーブル |
| DataTable | マスタデータテーブル |
| SpreadsheetTable | 自由形式の計算シート |

## インストール

### 前提条件

- [uv](https://docs.astral.sh/uv/) がインストール済みであること
- [Claude Code](https://claude.ai/code) がインストール済みであること（スキルを使う場合）

### インストール手順

**Windows (PowerShell):**
```powershell
git clone https://github.com/SzTk/openl-tablets-tool
cd openl-tablets-tool
.\install.ps1
```

**Mac / Linux (bash):**
```bash
git clone https://github.com/SzTk/openl-tablets-tool
cd openl-tablets-tool
bash install.sh
```

インストールスクリプトは以下を実行します：
1. `openl` コマンドをグローバルインストール (`uv tool install`)
2. Claude Code スキルを `~/.claude/skills/openl-edit/` に配置

### 手動インストール

```bash
# CLI ツールのみ
uv tool install .

# スキルのみ
mkdir -p ~/.claude/skills/openl-edit
cp skills/openl-edit/SKILL.md ~/.claude/skills/openl-edit/SKILL.md
```

## CLI の使い方

```bash
# Excel → JSON（編集用）
openl read TicketsPrice.xlsx

# Excel → YAML
openl read TicketsPrice.xlsx --format yaml

# JSON/YAML → Excel（書き戻し）
openl write TicketsPrice.json

# 出力先を指定
openl write TicketsPrice.json --out TicketsPrice_updated.xlsx

# ラウンドトリップ確認
openl roundtrip TicketsPrice.xlsx
```

## Claude Code スキルの使い方

Claude Code のチャットで以下のように呼び出します：

```
/openl-edit TicketsPrice.xlsx
```

スキルが以下を自動で行います：
1. Excel → JSON 変換
2. テーブル構造の表示
3. 編集指示の受け付け（自然言語で指定可能）
4. JSON 編集
5. JSON → Excel 書き戻し

**編集の例:**
```
> /openl-edit TicketsPrice.xlsx
(構造が表示される)

> NGO→SIN エコノミー 大人 を 35,000円で追加して
> 燃油サーチャージ率を10%に変更して
```

## ディレクトリ構成

```
openl-tablets-tool/
├── openl/              # Python パッケージ
│   ├── models.py       # Pydantic データモデル
│   ├── reader.py       # Excel → モデル パーサ
│   ├── writer.py       # モデル → Excel ライタ
│   └── cli.py          # CLI エントリポイント
├── skills/
│   └── openl-edit/
│       └── SKILL.md    # Claude Code スキル定義
├── install.sh          # インストールスクリプト (Mac/Linux)
├── install.ps1         # インストールスクリプト (Windows)
├── pyproject.toml      # プロジェクト設定
└── README.md
```

## ライセンス

MIT
