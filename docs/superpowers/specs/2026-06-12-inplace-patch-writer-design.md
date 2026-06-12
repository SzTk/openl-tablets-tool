# OpenL Tablets In-Place Patch Writer — Design Spec

**Date:** 2026-06-12
**Status:** Approved

## Context

`openl-tablets-edit` の現行フローは以下の通り:

```
openl read  <file.xlsx>  → <file>.json   (OpenLReader)
(JSONをLLMが編集)
openl write <file>.json  → <out>.xlsx    (OpenLWriter: ゼロから再構築)
```

`OpenLWriter` は JSON モデルからワークブックを**ゼロから再構築**するため、編集前後で「データの意味は同じだがフォーマットが変わる」問題が発生する。`examples/AutoPolicyCalculation.xlsx` で条件値を1箇所変更しただけのラウンドトリップで確認したところ:

- セルのテーマカラー（`theme`+`tint`）が読み取り時点で失われる（`_read_sheet_styles`はRGB色のみ対応）
- `sheet_styles` に記録された書式（フォント・塗り・罫線・数値書式）は `_apply_sheet_styles()` が**一度も呼ばれず**書き戻し時に丸ごと破棄される
- テーブルとして認識されないセル（シートタイトル等）はJSONに一切含まれず消失する
- `SpreadsheetTable`の書き込みが`start_col`を無視しB列固定になり、テーブル位置がズレる
- 結果としてシートの使用範囲が `B1:J65 → B1:D45` のように大幅に縮小する

この問題は初回実装時から認識されており（コミット `74a3605`）、当時は「リビルド後のレイアウトと記録した座標がズレる」ため `_apply_sheet_styles` を無効化する対症療法で済ませた経緯がある。

**根本原因:** 書き込み方式が「JSONからゼロから再構築」であること。再構築する限り、JSONモデルが表現しない情報（テーマ色・タイトルセル・列構成の暗黙の前提など）は原理的に失われる。

## Goals

- `openl-tablets-edit`での編集（値変更・行追加・行削除）について、**編集箇所以外のセルの書式・レイアウト・テーマ色・タイトル等を完全に保持**したまま書き戻す
- JSONスキーマをLLMにとってシンプルに保つ（内部座標情報等で汚染しない）
- `openl-tablets-create`（新規作成）の既存フローには影響を与えない

## Non-Goals

- ピクセル単位の完全なバイナリ一致（zip内部のXML再シリアライズによる差異は許容する）
- 条件列・結果列などテーブルの列構成変更時の書式保持（フォールバックで動作はするが書式は引き継がれない）
- マージセル・条件付き書式・コメント・データ検証等、現行モデルが扱っていない要素への対応

## アーキテクチャ

```
openl read <file.xlsx>
  → OpenLWorkbook (JSON)   ※ sheet_styles / sheet_dimensions は廃止
  → <file>.json として出力 (source_file: "<file.xlsx>" を保持)

openl write <edited.json> [--out <path>] [--source <original.xlsx>]
  1. 元ファイルを解決:
       --source 指定があればそれを使用
       なければ <edited.jsonと同じディレクトリ>/<source_file>
  2a. 解決パスが存在する     → パッチモード (patch_writer.py)
  2b. 解決パスが存在しない   → フルリビルドモード (既存 OpenLWriter、create フロー等)
       ※ 編集フローで元ファイルが移動/削除された場合もここにフォールバックし、
          「元ファイルが見つからないためフルリビルドします」と表示する
```

`openl-tablets-edit` / `openl-tablets-create` 両スキルのCLI呼び出し方は変更不要。`--source`は任意のオーバーライド用オプション。

### パッチモードの処理

```
a. 元ファイルを内部リーダー (read_with_positions) で再パース
   → 各テーブルの "before" 行リスト（Excel行番号つき）
b. 編集後JSON → "after" 行リスト
c. テーブルを (sheet_name, table_kind, テーブル識別名) で before/after 突き合わせ
   - SimpleDecisionTable / DataTable: `table_name`
   - SpreadsheetTable: `description`からメソッド名を抽出した値
     （`SimpleDecisionTable.table_name`と同じ抽出規則: シグネチャの3単語目の
     `(`より前。同一シートに複数のSpreadsheetTableが並ぶケースを区別するため）
d. 一致したテーブル → 行レベルdiff＆パッチ（下記）
e. before側のみに存在するテーブル → その行範囲を削除
f. after側のみに存在するテーブル → 既存writer関数でシート末尾/新規シートに追記
g. 元ワークブック(openpyxlオブジェクト)に対し上記操作を適用し --out に保存
```

## 行レベルdiff＆パッチアルゴリズム

### 比較用タプル（`id`や行番号は使わず内容のみで比較）

| テーブル種別 | 比較対象（1行分） |
|---|---|
| SimpleDecisionTable | `(条件値の並び..., 結果値の並び..., notes)` |
| DataTable | `columns`順に並べた `data` の値 |
| SpreadsheetTable | ステップ: `(label, value, unit)` / notes: 文字列 |

### 差分とパッチ操作

`difflib.SequenceMatcher(before_tuples, after_tuples)` で `equal / replace / insert / delete` を取得し、以下に変換する。**シート内で行番号の大きい方から順に適用**し、insert/deleteによる行ズレを回避する。

- `equal` → 何もしない
- `replace`（同行数）→ 値が変わったセルのみ `cell.value` を書き換え。スタイルは無変更
- `replace`（行数差あり）→ 共通部分は値のみ書き換え、余剰分をinsert/deleteとして処理
- `insert` → 直前行（無ければテーブル先頭データ行）をテンプレートとし `ws.insert_rows()` 後、各セルの `font/fill/border/number_format` をテンプレート行からコピーしてから値を設定
- `delete` → `ws.delete_rows()`（残存行のスタイルはopenpyxlが自動シフト）

### `"="`始まりの値（OpenL式）

既存のリビルドwriterと同様、`cell._value = val; cell.data_type = "s"` として明示的にテキスト型で書き込み、Excel数式として誤評価されるのを防ぐ。

### ヘッダー行・宣言行

シグネチャ行・列名行・`Datatype`宣言行はdiff対象外。列構成（`conditions`/`results`/`columns`/`column_names`）が一致している限りそのまま保持する。

## 新規テーブル・シート / 列構成変更時の扱い

### 新規テーブル（既存シートに新しい`tables[]`要素）

`(sheet_name, table_kind, table_name)`がbefore側に存在しない場合、シート末尾（`ws.max_row + 1`）に空行を挟み、既存の `_write_simple_decision_table` / `_write_data_table` / `_write_spreadsheet_table` をそのまま呼び出して追記する。新規部分のため元書式は無く、現行createフローと同じ見た目（openpyxlデフォルト）になる。

### 新規シート

`sheet_name`がbefore側のworkbookに存在しない場合、`wb.create_sheet(title=...)`で作成し、上記と同じ書き込み関数で書く。

### テーブル削除（before側のみに存在）

該当テーブルの行範囲（シグネチャ行〜最終データ行、区切りの空行含む）を`ws.delete_rows()`で削除する。

### 列構成変更（`conditions`/`results`/`columns`/`column_names`がbefore/afterで異なる）

行のinsert/deleteでは表現できないため、**そのテーブルのみ**「同じ位置で削除＋再作成」とする:
- 元の行範囲（ヘッダー行含む）を`ws.delete_rows()`で削除
- 同じ位置に`ws.insert_rows()`で空行を確保し、リビルドwriter関数で新しい列構成を書き込む

このテーブルのみ書式はopenpyxlデフォルトに戻る（他のテーブル・セルには影響しない）。これは稀なケース（条件列の追加等）であり、SKILL.mdの主要編集パターン（ルール追加・行追加・定数変更・パラメータ変更）には含まれない。

## スキーマ変更（`sheet_styles` / `sheet_dimensions` 廃止）

in-placeパッチ方式では触らないセルの書式は元ファイルにそのまま残るため、これらのフィールドは不要になる。

**削除対象:**
- `models.py`: `CellStyle`, `SheetDimensions`, `OpenLWorkbook.sheet_styles`, `OpenLWorkbook.sheet_dimensions`, `_DEFAULT_FONT_NAME`, `_DEFAULT_FONT_SIZE`
- `reader.py`: `_read_sheet_styles`, `_read_sheet_dimensions`, `_DEFAULT_FONT_COLORS`, `_DEFAULT_FILL_COLORS`, `_border_style`、および`OpenLReader.read()`内の呼び出し
- `writer.py`: 既に未使用の`_apply_sheet_styles`、および`_apply_sheet_dimensions`とその呼び出し
- `skills/openl-lib/SCHEMA.md`: スキーマ例から両フィールドを削除し、関連する「編集時は変更しない」の注記を削除
- `skills/openl-tablets-edit/SKILL.md`: 同様の注記を削除

`examples/AutoPolicyCalculation.json`は約1MB→数十KB程度に縮小する見込み。

## モジュール構成

**新規モジュール `openl/patch_writer.py`**
- `patch_write(edited: OpenLWorkbook, original_path: Path, out_path: Path) -> None`
- `reader.py`の`_parse_simple_rules` / `_parse_datatype` / `_parse_spreadsheet`を拡張し、各データ行の**Excel行番号リスト**も返すようにする。`OpenLReader.read()`は従来通りこれを破棄。新たに`read_with_positions(path)`を追加し、`patch_writer`が内部的に使用する
- `cli.py`の`cmd_write`で「元ファイル解決→存在チェック」を行い、`patch_write` / 既存`OpenLWriter.write`を振り分ける

## テスト方針

`openl`パッケージ向けの初のテストスイート（`tests/`）を新設する。「保持される」の判定基準は**openpyxlで読んだ際のセル値・フォント・塗り・罫線・数値書式・列幅行高の一致**（zipバイナリ完全一致ではない）。

1. ルールの条件値を1つ変更 → 該当セルのみ変化、他は全セル値・スタイルが一致
2. SimpleDecisionTable末尾にルール追加 → 新規行が直前行のスタイルをコピーして挿入される
3. ルールを途中に追加 → 正しい位置に挿入、以降の行のスタイル保持
4. ルールを削除 → 行削除、残りのスタイル保持
5. DataTableに行追加 / SpreadsheetTableにステップ追加
6. 既存シートに新規テーブル追加 / 新規シート追加
7. 列構成変更（条件列追加など）→ クラッシュせずフォールバック動作
8. **`examples/AutoPolicyCalculation.xlsx`を使った統合テスト**: 本設計のきっかけとなった編集（Accidents条件 `>2`→`>=10`）を適用し、(a) 値が正しく変わる、(b) タイトルセルとそのテーマカラー・フォントが保持される、(c) 編集対象外のシート・列幅・行高が変化しない、ことを確認
9. 元ファイルが見つからない場合のフルリビルドへのフォールバック（既存`ShopPolicy.xlsx`等の既存roundtripの回帰確認）

## 既知の制限・将来課題

- 列構成変更時はテーブル単位で書式が失われる（上記フォールバック）
- マージセル・条件付き書式・コメント・データ検証は引き続き非対応
- シート自体の削除（テーブルが0件になったシート）への対応は本設計の範囲外
