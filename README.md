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
2. Claude Code スキルを `~/.claude/skills/` に配置
   - `openl-tablets-edit/` — 既存 Excel の編集スキル
   - `openl-tablets-create/` — 新規作成スキル
   - `openl-lib/` — 両スキルが参照する共有スキーマ定義

### 手動インストール

```bash
# CLI ツールのみ
uv tool install .

# スキルのみ
mkdir -p ~/.claude/skills/openl-tablets-edit ~/.claude/skills/openl-tablets-create ~/.claude/skills/openl-lib
cp skills/openl-tablets-edit/SKILL.md ~/.claude/skills/openl-tablets-edit/SKILL.md
cp skills/openl-tablets-create/SKILL.md ~/.claude/skills/openl-tablets-create/SKILL.md
cp skills/openl-lib/SCHEMA.md ~/.claude/skills/openl-lib/SCHEMA.md
```

## CLI の使い方

```bash
# Excel → JSON（編集用）
openl read MyRules.xlsx

# Excel → YAML
openl read MyRules.xlsx --format yaml

# JSON/YAML → Excel（書き戻し）
openl write MyRules.json

# 出力先を指定
openl write MyRules.json --out MyRules_updated.xlsx

# ラウンドトリップ確認
openl roundtrip MyRules.xlsx
```

## Claude Code スキルの使い方

### openl-tablets-edit — 既存ファイルの編集

```
/openl-tablets-edit MyRules.xlsx
```

スキルが以下を自動で行います：
1. Excel → JSON 変換
2. テーブル構造の表示
3. 編集指示の受け付け（自然言語で指定可能）
4. JSON 編集
5. JSON → Excel 書き戻し

```
> /openl-tablets-edit MyRules.xlsx
(構造が表示される)

> NGO→SIN エコノミー 大人 を 35,000円で追加して
> 燃油サーチャージ率を10%に変更して
```

### openl-tablets-create — 新規ファイルの作成

```
/openl-tablets-create MyNewRules.json
```

以下の 3 層構造を対話的に設計します：

```
SpreadsheetTable    ← API エントリポイント（出力項目を定義）
      ↓ メソッド呼び出し
SimpleDecisionTable ← 各出力項目の計算ロジック（条件→結果ルール）
      ↓ 値域の参照
DataTable           ← マスタデータ・定数
```

ヒアリングの流れ：
1. **API 出力項目** — SpreadsheetTable の steps（出力項目名・型・呼び出すメソッド）
2. **計算ロジック** — 各 step が呼ぶ SimpleDecisionTable の条件・結果・ルール
3. **マスタデータ** — Claude が DataTable 候補を提案、ユーザーが採否を決定

最後に **往復検証**（JSON → Excel → JSON）を実行し、生成した JSON が正しく変換できることを確認します。

## ディレクトリ構成

```
openl-tablets-tool/
├── openl/              # Python パッケージ
│   ├── models.py       # Pydantic データモデル
│   ├── reader.py       # Excel → モデル パーサ
│   ├── writer.py       # モデル → Excel ライタ
│   └── cli.py          # CLI エントリポイント
├── skills/
│   ├── openl-tablets-edit/
│   │   └── SKILL.md    # 既存 Excel 編集スキル
│   ├── openl-tablets-create/
│   │   └── SKILL.md    # 新規作成スキル
│   └── openl-lib/
│       └── SCHEMA.md   # 共有 JSON スキーマ定義
├── install.sh          # インストールスクリプト (Mac/Linux)
├── install.ps1         # インストールスクリプト (Windows)
├── pyproject.toml      # プロジェクト設定
└── README.md
```

## ライセンス

MIT
