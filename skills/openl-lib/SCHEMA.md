# OpenL Tablets JSON スキーマ共有リファレンス

`openl-tablets-edit` / `openl-tablets-create` 両スキルから参照する共通定義。

## テーブル種別の役割と関係

```
API呼び出し
    ↓
SpreadsheetTable          ← API エントリポイント。出力項目を step として定義
  step.value = "= MethodName(params)"
                ↓
  SimpleDecisionTable     ← 各 step の計算ロジック。条件→結果のルールセット
    conditions に使う値の定義域
                ↓
  DataTable               ← マスタデータ / 定数。複数テーブルをまたいで参照
```

## JSON スキーマ早見表

```jsonc
{
  "source_file": "MyRules.xlsx",
  "tables": [

    // ── SpreadsheetTable ──────────────────────────────────────────────
    // API エントリポイント。parameters が入力、steps が出力項目。
    // step.value に "= MethodName(args)" を書いて SimpleDecisionTable を呼ぶ。
    // 他の step を参照するときは "$StepLabel" を使う。
    // column_names: Excel に出力される列ヘッダー。省略時は ["Step", "Description", "Value"]。
    //               通常は ["Step", "Value"] を使う（2列）。
    {
      "table_kind": "SpreadsheetTable",
      "sheet_name": "Calculation",
      "title": "",
      "description": "Spreadsheet SpreadsheetResult MethodName (TypeA paramA, TypeB paramB)",
      "parameters": [],          // parameters は description のシグネチャで表現
      "column_names": ["Step", "Value"],
      "steps": [
        { "label": "InputParam",  "value": "= paramA"              },
        { "label": "StepA",       "value": "= CalcMethodA(paramA, paramB)"   },
        { "label": "StepB",       "value": "= CalcMethodB($StepA)"           },
        { "label": "FinalResult", "value": "= $StepA + $StepB"               }
      ],
      "start_col": 2,
      "notes": []
    },

    // ── SimpleDecisionTable ───────────────────────────────────────────
    // 条件→結果のルールテーブル。
    // method_signature の MethodName が SpreadsheetTable の step.value 呼び出し名と一致すること。
    // String 型の条件値はそのまま文字列で書く（例: "Tokyo"）。
    // int の範囲は "18 .. 64" 形式（両端を含む）または ">= 18"、"< 65" 形式。
    // null 条件 = ワイルドカード（どの値にもマッチ）。
    {
      "table_kind": "SimpleDecisionTable",
      "sheet_name": "RulesSheet",
      "title": "",
      "method_signature": "SimpleRules ReturnType MethodName (TypeA paramA, TypeB paramB)",
      "table_name": "MethodName",
      "conditions": [
        { "name": "ParamA Label", "col_type": "String", "role": "condition" },
        { "name": "ParamB Label", "col_type": "double", "role": "condition" }
      ],
      "results": [
        { "name": "Result Label", "col_type": "String", "role": "result" }
      ],
      "rules": [
        {
          "id": 1,
          "conditions": { "ParamA Label": "ValueX", "ParamB Label": ">= 100" },
          "results":    { "Result Label": "High" },
          "notes": null
        }
      ],
      "start_col": 1
    },

    // ── DataTable ─────────────────────────────────────────────────────
    // マスタデータ / 定数テーブル。
    // String 型の値はそのまま文字列。数値はそのまま。
    {
      "table_kind": "DataTable",
      "sheet_name": "Constants",
      "title": "",
      "table_type": "Constants",
      "table_name": "ConstantValue",
      "columns": [
        { "name": "Name",  "col_type": "String", "role": "data" },
        { "name": "Value", "col_type": "double", "role": "data" }
      ],
      "rows": [
        { "data": { "Name": "RATE_A", "Value": 0.05 } }
      ],
      "start_col": 1
    }

  ]
}
```

## 型と値の記法

| col_type | 条件値の例 | 結果値の例 |
|----------|-----------|-----------|
| String   | `"Tokyo"` | `"High"` |
| int      | `">= 18"` / `"18 .. 64"` | `42` |
| double   | `"< 100000"` | `0.08` |
| Boolean  | `"true"` | `true` |

## method_signature の形式

```
SimpleRules <ReturnType> <MethodName> (<Type1> <param1>, <Type2> <param2>)
```

- `ReturnType`: SimpleDecisionTable の結果列の型（例: `String`, `double`, `TheftRating`）
- `MethodName`: SpreadsheetTable の step.value から呼び出す名前と一致させること
- 引数名は SpreadsheetTable の description シグネチャの引数名またはフィールド名と一致させること

## 制約

- `id` は 1 始まりの連番、途中を欠番にしない
- `start_col` は通常 1（SpreadsheetTable は 2）
