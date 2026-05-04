---
name: openl-edit
description: OpenL Tablets 形式の Excel ファイルを読み取り、内容を確認・編集して書き戻す。OpenL Tablets の SimpleDecisionTable / DataTable / SpreadsheetTable を含む Excel を編集する場合に使う。
license: MIT
metadata:
  argument-hint: "<excel_file_path>"
allowed-tools: Bash Read Write Edit
---

# OpenL Tablets 編集スキル

OpenL Tablets 形式の Excel ファイルを読み書きする `openl` CLI を使い、確認・編集・書き戻しを行う。

## 前提

`openl` コマンドがインストール済みであること。
未インストールの場合は以下を実行:
```bash
uv tool install git+https://github.com/SzTk/openl-tablets-tool
```
または
```bash
cd <リポジトリ>/excel-tool && uv tool install .
```

## ステップ

### Step 1: 対象ファイルの確定

引数 `$ARGUMENTS` が指定されていればそれを対象ファイルとする。
指定がなければ、カレントディレクトリの `.xlsx` ファイルを一覧してユーザーに選択を求める。

対象ファイルのパスを絶対パスに解決する。

### Step 2: Excel → JSON 変換

```bash
openl read <対象Excelの絶対パス> --format json
```

JSON は Excel と同じディレクトリに `<元ファイル名>.json` として出力される。

### Step 3: 構造サマリーの表示

JSON を Read ツールで読み取り、以下の形式でユーザーに示す。

```
ファイル: <ファイル名>

[PriceRules] SimpleDecisionTable
  条件: Origin (String), Destination (String), Class (String), Age (int)
  結果: BasePrice (double)
  ルール数: 12 件

[DiscountRules] SimpleDecisionTable
  ...

[Airports] DataTable
  列: Code, City, Country, Region, Notes
  行数: 12 件

[Constants] DataTable
  ...

[FinalPriceCalc] SpreadsheetTable
  パラメータ: 6 件 / ステップ: 6 件
```

### Step 4: 編集内容の受け付け

ユーザーの指示に従い、JSON ファイルを Edit/Write ツールで変更する。

**編集の典型例:**

- **ルール追加**: `tables[n].rules` に新しい `Rule` オブジェクトを追加。`id` は末尾の最大値 +1。
- **ルール変更**: 対象ルールの `conditions` または `results` の値を変更。
- **行追加 (DataTable)**: `tables[n].rows` に `{"data": {...}}` を追加。
- **定数変更 (Constants)**: 対象行の `data.Value` を変更。
- **パラメータ変更 (SpreadsheetTable)**: `parameters` の `value` を変更。

**制約の遵守:**
- `id` の連番を崩さない
- 型 (`col_type`) に合った値を使う（String は `"\"値\""` 形式、数値はそのまま）
- `conditions` の書式（`>= 12 && < 65` など）を既存ルールに合わせる

### Step 5: JSON → Excel 書き戻し

```bash
openl write <JSONファイルパス> --out <元ファイル名>_edited.xlsx
```

元ファイルを上書きする場合はユーザーの明示的な確認を得てから `--out <元のExcelパス>` を使う。

### Step 6: 完了報告

```
✅ 編集完了

変更内容:
  - <変更点の要約>

出力ファイル: <出力パス>
```

## JSON スキーマ

JSON 構造・型記法・制約の詳細は `skills/openl-lib/SCHEMA.md` を Read して参照すること。

## 注意事項

- `sheet_styles` と `sheet_dimensions` はルール編集時は変更しない
- String 型の条件値は OpenL 記法に従い `"\"値\""` 形式でラップする
