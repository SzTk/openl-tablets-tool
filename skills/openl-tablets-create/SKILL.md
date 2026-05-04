---
name: openl-tablets-create
description: OpenL Tablets 形式の Excel ファイルを対話的に新規作成する。API出力項目→計算ロジック→マスタデータの順でヒアリングし、JSON を生成・往復検証して Excel を出力する。
license: MIT
metadata:
  argument-hint: "[output_file_name]"
allowed-tools: Bash Read Write Edit
---

# OpenL Tablets 新規作成スキル

`openl` CLI を使って OpenL Tablets 形式の JSON/Excel ファイルをゼロから対話的に作成する。

## 前提

`openl` コマンドがインストール済みであること。
未インストールの場合は以下を実行:
```bash
uv tool install git+https://github.com/SzTk/openl-tablets-tool
```

**スキーマ参照**: 作業前に `skills/openl-lib/SCHEMA.md` を Read して JSON 構造・型記法・制約を確認すること。

## テーブル設計の考え方

このスキルは以下の 3 層構造で OpenL Tablets ファイルを設計する。

```
SpreadsheetTable   ← API エントリポイント。出力項目（steps）を定義
      ↓ step.value = "= MethodName(args)" で呼び出す
SimpleDecisionTable ← 各出力項目の計算ロジック（条件→結果ルール）
      ↓ 条件の値域を定義
DataTable           ← マスタデータ・定数（複数テーブルをまたいで共有）
```

## ステップ

### Step 1-A: API 出力項目のヒアリング（SpreadsheetTable 設計）

ユーザーに以下を確認する:

1. **出力ファイル名**（引数 `$ARGUMENTS` が指定されていればそれを使う。なければ確認する）
2. **API 関数名**: SpreadsheetTable のメソッド名（例: `DetermineVehiclePremium`）
3. **入力パラメータ**: 型と名前（例: `Vehicle vehicle`, `int age`）
4. **出力項目（steps）の一覧**: ラベル・説明・呼び出すメソッド名（後で確定してもよい）
5. **SpreadsheetTable の列名**: Excel に出力される列ヘッダー。デフォルト `["Step", "Value"]` を提案し、変更する場合のみ聞く。

確認結果をまとめて表示し、ユーザーの承認を得る:

```
SpreadsheetTable: DetermineVehiclePremium (Vehicle vehicle)
  列名: ["Step", "Value"]
  step[1]: TheftRating   ← VehicleTheftRating(bodyType, price, ...) を呼ぶ予定
  step[2]: InjuryRating  ← VehicleInjuryRating(bodyType, airbagType) を呼ぶ予定
  step[3]: Premium       ← 最終結果（計算式）
```

### Step 1-B: 各出力項目の計算ロジックヒアリング（SimpleDecisionTable 設計）

Step 1-A で確定した steps のうち、SimpleDecisionTable を呼ぶもの（`= MethodName(...)` 形式）を 1 件ずつヒアリングする。

各 SimpleDecisionTable について確認する内容:
- **メソッド名と戻り値型**（例: `TheftRating VehicleTheftRating`）
- **引数（条件列）**: 名前・型・表示ラベル
- **結果列**: 名前・型
- **ルール**: 条件の組み合わせと対応する結果値

```
[VehicleTheftRating] の条件を教えてください:
  条件列: bodyType (String), price (double), onHighTheftProbabilityList (Boolean)
  結果列: TheftRating (String)

  ルールを入力してください（例: bodyType=Convertible → High）:
  → 入力が終わったら「完了」と言ってください
```

すべての SimpleDecisionTable のヒアリングが終わったら、SpreadsheetTable の `step.value` 式を確定させる（`$StepLabel` 参照を含む）。

### Step 1-C: DataTable の提案（Claude が主導）

Step 1-B の内容から Claude が以下を分析して DataTable 候補を提案する:

**提案の観点:**
- 複数の SimpleDecisionTable で同じ String 型の値域が使われている列
- ルール数が多く、値の定義を外部化すると保守しやすいもの
- 定数（固定の数値・文字列）をまとめられるもの

**提案フォーマット:**
```
以下の DataTable を提案します:

[BodyTypeList] DataTable
  列: Code (String), Label (String)
  理由: VehicleTheftRating と VehicleInjuryRating の両方で bodyType の条件として使用
  → 採用 / 不要 / 変更

[Constants] DataTable
  列: Name (String), Value (double)
  候補定数: BASE_PREMIUM_RATE = 0.05, DISCOUNT_MAX = 0.30
  → 採用 / 不要 / 変更
```

ユーザーの採否を確認して DataTable 一覧を確定させる。

### Step 2: テーブル全体構造の確認・合意

すべてのテーブルをまとめてサマリー表示し、最終確認を得る:

```
生成するテーブル構成:

[Calculation] SpreadsheetTable
  関数: DetermineVehiclePremium (Vehicle vehicle)
  steps: 8 件

[Vehicle-Eligibility] SimpleDecisionTable × 3
  VehicleTheftRating  / VehicleInjuryRating / VehicleEligibility

[MasterData] DataTable × 2
  BodyTypeList / Constants

シート名・テーブル名に問題がなければ「OK」と言ってください。
```

### Step 3: JSON 骨格生成

ユーザーの承認後、Write ツールで `<ファイル名>.json` を生成する。

- `source_file` は `<ファイル名>.xlsx`
- `sheet_styles`: `{}`、`sheet_dimensions`: `{}`
- ルールは Step 1-B で確定した内容をすべて含める
- `id` は 1 始まり連番
- SpreadsheetTable に `column_names` を設定する（Step 1-A で確定した値。デフォルトは `["Step", "Value"]`）

### Step 4: ルール・データの追加編集

ユーザーから追加のルール・行・修正があれば Edit/Write ツールで JSON を更新する。
「検証に進む」または「完了」と言われるまでこのフェーズを継続する。

### Step 5: 往復検証

JSON が正しく Excel に変換でき、再び JSON に戻せることを確認する。

```bash
# JSON → Excel
openl write <file>.json --out <file>_draft.xlsx

# Excel → JSON（再読み込み）
openl read <file>_draft.xlsx --out <file>_rt.json
```

Read ツールで `<file>_rt.json` を読み込み、以下の点を元の JSON と比較する:

| 確認項目 | 合格条件 |
|----------|---------|
| テーブル数・種別 | 完全一致 |
| 各テーブルのカラム定義 | 完全一致 |
| SimpleDecisionTable のルール数 | 完全一致 |
| DataTable の行数 | 完全一致 |
| SpreadsheetTable の step 数 | 完全一致 |
| 代表ルール値のサンプル（3 件） | 値が一致 |

**不一致が検出された場合:**
- 差分をユーザーに報告する
- JSON の修正が必要な箇所を特定して修正案を提示する
- 修正後に往復検証を再実行する

**合格した場合:**
```
往復検証: OK
  テーブル数: 6 / 6 一致
  ルール総数: 34 / 34 一致
```

### Step 6: 完了報告

```
✅ 新規作成完了

テーブル構成:
  [Calculation]        SpreadsheetTable   steps: 8
  [VehicleTheftRating] SimpleDecisionTable ルール数: 5
  [VehicleInjuryRating] SimpleDecisionTable ルール数: 6
  [Constants]          DataTable          行数: 3

往復検証: OK

出力 JSON : MyRules.json
出力 Excel: MyRules_draft.xlsx
```

元 Excel ファイルとして `_draft` なしの名前で保存したい場合はユーザーに確認してから:
```bash
openl write <file>.json --out <file>.xlsx
```
