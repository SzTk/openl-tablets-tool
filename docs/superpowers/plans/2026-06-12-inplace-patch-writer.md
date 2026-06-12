# In-Place Patch Writer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `openl write` を、編集箇所だけを元の Excel ファイルに適用する「パッチモード」に対応させ、編集対象外のセルの書式・レイアウト・テーマ色を完全に保持する。

**Architecture:** 新規モジュール `openl/patch_writer.py` に `patch_write(edited, original_path, out_path)` を実装する。`read_with_positions()`（既存）で元ファイルを再パースし、`(sheet_name, table_kind, identifier)` でテーブルをbefore/after突き合わせ、`difflib.SequenceMatcher` による行レベルdiffを `equal/replace/insert/delete` パッチ操作に変換してopenpyxlワークブックへ適用する。新規/削除テーブル・列構成変更は別経路で処理する。`cli.py` の `cmd_write` は元ファイルの有無でパッチモード/フルリビルドモードを振り分ける。

**Tech Stack:** Python 3.12, openpyxl, pydantic v2, pytest, difflib (標準ライブラリ)

設計の詳細・背景は `docs/superpowers/specs/2026-06-12-inplace-patch-writer-design.md`（Approved）を参照。

---

## File Structure

- Create: `openl/patch_writer.py` — `patch_write()` とテーブル突き合わせ・行diff・パッチ適用のヘルパー一式
- Create: `tests/test_patch_writer.py` — Task 1〜5・統合テスト
- Modify: `openl/cli.py:65-78`（`cmd_write`）— パッチモード/フルリビルドモードの振り分け、`--source` オプション追加
- Test (CLI): `tests/test_cli_write.py` — `cmd_write` の振り分けロジック

## 前提知識（既存コードの再利用箇所）

- `openl/reader.py` の `read_with_positions(path) -> list[ParsedTable]`、`ParsedTable(table, start_row, header_rows, item_rows, end_row)`（すべて1-based Excel行番号）
- `openl/writer.py` の `_write_simple_decision_table(ws, table)`, `_write_data_table(ws, table)`, `_write_spreadsheet_table(ws, table)` — 新規テーブル追記・再作成に再利用する
- `OpenLWorkbook.source_file: str`（既存フィールド、元ファイル名を保持）

---

### Task 1: テーブル識別・行比較ヘルパー

**Files:**
- Create: `openl/patch_writer.py`
- Create: `tests/test_patch_writer.py`

これから実装する `patch_write()` がbefore/afterのテーブル・行を突き合わせるための純粋関数群。I/Oなし。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_patch_writer.py` を新規作成:

```python
from pathlib import Path

import openpyxl

from openl.reader import OpenLReader
from openl.patch_writer import (
    _table_identity,
    _structure_key,
    _comparison_tuples,
    _row_cell_values,
    _set_cell,
)

SHOP_POLICY = Path(__file__).parent.parent / "examples" / "ShopPolicy.xlsx"
AUTO_POLICY = Path(__file__).parent.parent / "examples" / "AutoPolicyCalculation.xlsx"


def _shop_policy():
    return OpenLReader().read(SHOP_POLICY)


def test_table_identity_simple_decision_table():
    wb = _shop_policy()
    table = wb.get_table("FreeShipping")
    assert _table_identity(table) == ("FreeShipping", "SimpleDecisionTable", "IsFreeShipping")


def test_table_identity_spreadsheet_table():
    wb = _shop_policy()
    table = wb.get_table("Calculation")
    assert table.table_kind == "SpreadsheetTable"
    assert _table_identity(table) == ("Calculation", "SpreadsheetTable", "DetermineShopPolicy")


def test_structure_key_matches_for_unchanged_table():
    wb = _shop_policy()
    table = wb.get_table("FreeShipping")
    assert _structure_key(table) == _structure_key(table.model_copy(deep=True))


def test_structure_key_differs_when_condition_added():
    wb = _shop_policy()
    table = wb.get_table("FreeShipping")
    changed = table.model_copy(deep=True)
    from openl.models import ColumnDef
    changed.conditions.append(ColumnDef(name="新条件", col_type="String", role="condition"))
    assert _structure_key(table) != _structure_key(changed)


def test_comparison_tuples_simple_decision_table():
    wb = _shop_policy()
    table = wb.get_table("FreeShipping")
    tuples = _comparison_tuples(table)
    assert len(tuples) == 9
    assert tuples[0] == ("非会員", "< 3000", False, None)


def test_row_cell_values_simple_decision_table():
    wb = _shop_policy()
    table = wb.get_table("FreeShipping")
    assert _row_cell_values(table, 0) == ["非会員", "< 3000", False]


def test_comparison_tuples_spreadsheet_table():
    wb = _shop_policy()
    table = wb.get_table("Calculation")
    tuples = _comparison_tuples(table)
    assert tuples[0] == ("FreeShipping", "= IsFreeShipping(memberType, purchaseAmount)", "送料無料フラグ")


def test_set_cell_writes_formula_text_as_string():
    wb = openpyxl.Workbook()
    ws = wb.active
    _set_cell(ws, 1, 1, "= 1 + 1")
    cell = ws.cell(row=1, column=1)
    assert cell.value == "= 1 + 1"
    assert cell.data_type == "s"


def test_set_cell_writes_plain_value():
    wb = openpyxl.Workbook()
    ws = wb.active
    _set_cell(ws, 1, 1, 123)
    assert ws.cell(row=1, column=1).value == 123
```

- [ ] **Step 2: テストを実行して失敗を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py -v`
Expected: FAIL（`openl.patch_writer` モジュールが存在しないため `ModuleNotFoundError`）

- [ ] **Step 3: `openl/patch_writer.py` を実装する**

```python
"""
OpenL Tablets In-Place Patch Writer

編集後の OpenLWorkbook と元の Excel ファイルを比較し、変更箇所だけを
元ファイルに適用して書き戻す。編集対象外セルの書式・レイアウトは
そのまま保持される。
"""

from __future__ import annotations
from copy import copy
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from .models import AnyTable, OpenLWorkbook
from .reader import ParsedTable, read_with_positions
from .writer import _write_simple_decision_table, _write_data_table, _write_spreadsheet_table


TableIdentity = tuple[str, str, str]

_WRITER_MAP = {
    "SimpleDecisionTable": _write_simple_decision_table,
    "DataTable": _write_data_table,
    "SpreadsheetTable": _write_spreadsheet_table,
}


def _table_identity(table: AnyTable) -> TableIdentity:
    """(sheet_name, table_kind, identifier) でテーブルを一意に識別する。

    SpreadsheetTable は table_name を持たないため、description（シグネチャ）
    の3単語目から '(' より前を抽出する。SimpleRules/Spreadsheet 共通の
    シグネチャ書式 "<Keyword> <ReturnType> <Name>(<params>)" に依存する。
    """
    if table.table_kind == "SpreadsheetTable":
        parts = (table.description or "").split()
        ident = parts[2].split("(")[0] if len(parts) >= 3 else (table.description or "")
        return (table.sheet_name, table.table_kind, ident.strip())
    return (table.sheet_name, table.table_kind, table.table_name)


def _structure_key(table: AnyTable) -> tuple:
    """列構成の比較用キー。before/afterで異なれば delete+recreate の対象。"""
    if table.table_kind == "SimpleDecisionTable":
        return (
            tuple((c.name, c.col_type) for c in table.conditions),
            tuple((c.name, c.col_type) for c in table.results),
        )
    if table.table_kind == "DataTable":
        return tuple((c.name, c.col_type) for c in table.columns)
    return tuple(table.column_names)  # SpreadsheetTable


def _comparison_tuples(table: AnyTable) -> list[tuple]:
    """diff用の比較タプル列。ParsedTable.item_rows と要素数・順序が対応する。"""
    if table.table_kind == "SimpleDecisionTable":
        return [
            tuple(rule.conditions.get(c.name) for c in table.conditions)
            + tuple(rule.results.get(r.name) for r in table.results)
            + (rule.notes,)
            for rule in table.rules
        ]
    if table.table_kind == "DataTable":
        return [tuple(row.data.get(c.name) for c in table.columns) for row in table.rows]
    return [(step.label, step.value, step.unit) for step in table.steps]


def _row_cell_values(table: AnyTable, index: int) -> list[Any]:
    """item index 番目のデータ行を、start_col直後から並ぶセル値のリストとして返す。"""
    if table.table_kind == "SimpleDecisionTable":
        rule = table.rules[index]
        return [rule.conditions.get(c.name) for c in table.conditions] + \
               [rule.results.get(r.name) for r in table.results]
    if table.table_kind == "DataTable":
        row = table.rows[index]
        return [row.data.get(c.name) for c in table.columns]
    step = table.steps[index]
    if len(table.column_names) >= 3:
        return [step.label, step.unit, step.value]
    return [step.label, step.value]


def _set_cell(ws: Worksheet, row: int, col: int, value: Any) -> None:
    """セルに値を書き込む。'=' 始まりの文字列は OpenL 式としてテキスト型で書く。"""
    cell = ws.cell(row=row, column=col)
    if isinstance(value, str) and value.startswith("="):
        cell._value = value
        cell.data_type = "s"
    else:
        cell.value = value
```

- [ ] **Step 4: テストを実行して成功を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py -v`
Expected: `9 passed`

- [ ] **Step 5: コミット**

```bash
git add openl/patch_writer.py tests/test_patch_writer.py
git commit -m "feat: add table identity and row-comparison helpers for patch writer"
```

---

### Task 2: 行レベルdiff＆パッチ適用（値変更）と `patch_write` の骨格

**Files:**
- Modify: `openl/patch_writer.py`
- Modify: `tests/test_patch_writer.py`

`_comparison_tuples` を `difflib.SequenceMatcher` で比較し、`equal/replace/insert/delete` をセル書き込み・行挿入・行削除に変換する `_patch_table_rows` を実装する。あわせて、一致したテーブルのみを処理する `patch_write()` の骨格を作る（新規/削除テーブル・列構成変更は Task 4・5 で追加）。

このタスクのテストでは「値変更のみ（行数変化なし）」のシナリオ（design specのシナリオ1）を検証する。`_patch_table_rows` 自体は `insert`/`delete` opcode のハンドリングも含めて完全実装するが、それらは Task 3 で別シナリオから検証する。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_patch_writer.py` の末尾に追記:

```python
def _all_cell_values(wb):
    """openpyxl Workbook の全シート・全セル値を {sheet_name: [[...], ...]} で返す。"""
    return {
        name: [[c.value for c in row] for row in wb[name].iter_rows()]
        for name in wb.sheetnames
    }


def test_patch_write_value_change_preserves_everything_else(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(SHOP_POLICY)
    table = edited.get_table("FreeShipping")
    table.rules[0].results["送料無料"] = True  # was False

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, SHOP_POLICY, out_path)

    before_wb = openpyxl.load_workbook(SHOP_POLICY)
    after_wb = openpyxl.load_workbook(out_path)

    before = _all_cell_values(before_wb)
    after = _all_cell_values(after_wb)

    # 変更したセル（FreeShipping シート 3行目 D列 = 結果列）
    assert before["FreeShipping"][2][3] is False
    assert after["FreeShipping"][2][3] is True

    # 同じ行の他のセルは不変
    assert after["FreeShipping"][2][1:3] == before["FreeShipping"][2][1:3]

    # 変更したセル以外はすべて一致
    after["FreeShipping"][2][3] = before["FreeShipping"][2][3]
    assert after == before
```

- [ ] **Step 2: テストを実行して失敗を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py::test_patch_write_value_change_preserves_everything_else -v`
Expected: FAIL（`ImportError: cannot import name 'patch_write'`）

- [ ] **Step 3: `_patch_table_rows` と `patch_write` を実装する**

`openl/patch_writer.py` の末尾に追記:

```python
def _copy_row_style(ws: Worksheet, template_row: int, target_row: int, columns: range) -> None:
    """target_row の各セルに template_row のスタイルをコピーする。"""
    for col in columns:
        src = ws.cell(row=template_row, column=col)
        dst = ws.cell(row=target_row, column=col)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format


def _insert_table_rows(
    ws: Worksheet, after: AnyTable, j1: int, j2: int, anchor_row: int, start_col: int,
) -> None:
    """anchor_row の直後に after の item j1..j2 を挿入し、anchor_row のスタイルをコピーする。"""
    count = j2 - j1
    insert_at = anchor_row + 1
    ws.insert_rows(insert_at, count)
    n_cols = len(_row_cell_values(after, j1))
    columns = range(start_col, start_col + n_cols)
    for offset in range(count):
        row_num = insert_at + offset
        _copy_row_style(ws, anchor_row, row_num, columns)
        for col, val in zip(columns, _row_cell_values(after, j1 + offset)):
            _set_cell(ws, row_num, col, val)


def _patch_table_rows(ws: Worksheet, before: ParsedTable, after: AnyTable) -> None:
    """before/after の比較タプルを SequenceMatcher で diff し、行単位でパッチを適用する。

    行番号がずれないよう、opcode はシート内で行番号の大きい方から処理する
    （SequenceMatcher.get_opcodes() の逆順）。item_rows はテーブル内で
    連続した行番号であることを前提とする。
    """
    before_tuples = _comparison_tuples(before.table)
    after_tuples = _comparison_tuples(after)
    matcher = SequenceMatcher(a=before_tuples, b=after_tuples, autojunk=False)
    start_col = after.start_col + 1  # 1-based

    for tag, i1, i2, j1, j2 in reversed(matcher.get_opcodes()):
        if tag == "equal":
            continue

        if tag == "replace":
            common = min(i2 - i1, j2 - j1)
            for k in range(common):
                row_num = before.item_rows[i1 + k]
                for col, val in zip(
                    range(start_col, start_col + len(_row_cell_values(after, j1 + k))),
                    _row_cell_values(after, j1 + k),
                ):
                    _set_cell(ws, row_num, col, val)
            if i2 - i1 > common:
                rows = before.item_rows[i1 + common:i2]
                ws.delete_rows(rows[0], len(rows))
            elif j2 - j1 > common:
                anchor = before.item_rows[i1 + common - 1]
                _insert_table_rows(ws, after, j1 + common, j2, anchor, start_col)

        elif tag == "delete":
            rows = before.item_rows[i1:i2]
            ws.delete_rows(rows[0], len(rows))

        elif tag == "insert":
            if i1 > 0:
                anchor = before.item_rows[i1 - 1]
            elif before.item_rows:
                anchor = before.item_rows[0] - 1
            else:
                anchor = before.start_row + before.header_rows - 1
            _insert_table_rows(ws, after, j1, j2, anchor, start_col)


def patch_write(edited: OpenLWorkbook, original_path: str | Path, out_path: str | Path) -> None:
    """編集後の edited を、original_path の書式・レイアウトを保ったまま out_path に書き出す。"""
    wb = openpyxl.load_workbook(str(original_path))
    before_parsed = read_with_positions(original_path)

    before_by_id = {_table_identity(p.table): p for p in before_parsed}
    after_by_id = {_table_identity(t): t for t in edited.tables}

    for ident, parsed in before_by_id.items():
        after = after_by_id.get(ident)
        if after is None:
            continue
        ws = wb[ident[0]]
        _patch_table_rows(ws, parsed, after)

    wb.save(str(out_path))
```

- [ ] **Step 4: テストを実行して成功を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py -v`
Expected: `10 passed`

- [ ] **Step 5: コミット**

```bash
git add openl/patch_writer.py tests/test_patch_writer.py
git commit -m "feat: implement row-level diff/patch for value changes in patch_write"
```

---

### Task 3: 行追加・行挿入・行削除のテスト（スタイルコピー検証）

**Files:**
- Modify: `tests/test_patch_writer.py`

Task 2 で実装した `_patch_table_rows` の insert/delete 経路を、design specのシナリオ2〜5で検証する。バグがあればこのタスク内で `openl/patch_writer.py` を修正する。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_patch_writer.py` の末尾に追記:

```python
from openl.models import Rule, DataTableRow, SpreadsheetStep


def test_patch_write_append_rule_to_end(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(SHOP_POLICY)
    table = edited.get_table("FreeShipping")
    table.rules.append(Rule(
        id=10,
        conditions={"会員種別": "プレミアム会員", "購入金額": "special"},
        results={"送料無料": True},
    ))

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, SHOP_POLICY, out_path)

    wb = openpyxl.load_workbook(out_path)
    ws = wb["FreeShipping"]

    assert ws.max_row == 12
    assert [ws.cell(12, c).value for c in (2, 3, 4)] == ["プレミアム会員", "special", True]

    # 直前行（11行目）からスタイルがコピーされている
    for col in (2, 3, 4):
        assert ws.cell(12, col).font == ws.cell(11, col).font
        assert ws.cell(12, col).number_format == ws.cell(11, col).number_format

    # 既存行は不変
    before_wb = openpyxl.load_workbook(SHOP_POLICY)
    before_ws = before_wb["FreeShipping"]
    for r in range(1, 12):
        assert [ws.cell(r, c).value for c in (2, 3, 4)] == \
               [before_ws.cell(r, c).value for c in (2, 3, 4)]


def test_patch_write_insert_rule_in_middle(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(SHOP_POLICY)
    table = edited.get_table("FreeShipping")
    new_rule = Rule(
        id=99,
        conditions={"会員種別": "一般会員", "購入金額": "1 .. 99"},
        results={"送料無料": False},
    )
    table.rules.insert(2, new_rule)  # rule1, rule2 の後に挿入

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, SHOP_POLICY, out_path)

    wb = openpyxl.load_workbook(out_path)
    ws = wb["FreeShipping"]
    before_wb = openpyxl.load_workbook(SHOP_POLICY)
    before_ws = before_wb["FreeShipping"]

    assert ws.max_row == 12

    # 行3,4（rule1,rule2）は不変
    for r in (3, 4):
        assert [ws.cell(r, c).value for c in (2, 3, 4)] == \
               [before_ws.cell(r, c).value for c in (2, 3, 4)]

    # 行5 に新規ルールが挿入され、スタイルは行4からコピー
    assert [ws.cell(5, c).value for c in (2, 3, 4)] == ["一般会員", "1 .. 99", False]
    for col in (2, 3, 4):
        assert ws.cell(5, col).font == ws.cell(4, col).font

    # 行6〜12 に元rule3〜rule9がシフトしている
    for offset, r in enumerate(range(6, 13)):
        orig_r = 5 + offset
        assert [ws.cell(r, c).value for c in (2, 3, 4)] == \
               [before_ws.cell(orig_r, c).value for c in (2, 3, 4)]


def test_patch_write_delete_rule(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(SHOP_POLICY)
    table = edited.get_table("FreeShipping")
    del table.rules[4]  # rule id=5 (一般会員 / 3000..9999) を削除

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, SHOP_POLICY, out_path)

    wb = openpyxl.load_workbook(out_path)
    ws = wb["FreeShipping"]
    before_wb = openpyxl.load_workbook(SHOP_POLICY)
    before_ws = before_wb["FreeShipping"]

    assert ws.max_row == 10

    # 行3〜6（rule1〜4）は不変
    for r in range(3, 7):
        assert [ws.cell(r, c).value for c in (2, 3, 4)] == \
               [before_ws.cell(r, c).value for c in (2, 3, 4)]

    # 行7〜10 に元rule6〜rule9がシフトしている
    for offset, r in enumerate(range(7, 11)):
        orig_r = 8 + offset
        assert [ws.cell(r, c).value for c in (2, 3, 4)] == \
               [before_ws.cell(orig_r, c).value for c in (2, 3, 4)]


def test_patch_write_add_data_table_row(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(AUTO_POLICY)
    table = next(
        t for t in edited.tables
        if t.sheet_name == "Vocabulary" and getattr(t, "table_type", None) == "TheftRating"
    )
    table.rows.append(DataTableRow(data={"value": "VeryHigh"}))

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, AUTO_POLICY, out_path)

    wb = openpyxl.load_workbook(out_path)
    ws = wb["Vocabulary"]
    before_wb = openpyxl.load_workbook(AUTO_POLICY)
    before_ws = before_wb["Vocabulary"]

    # 新しい行が9行目（直前の8行目の直後）に挿入される
    assert ws.cell(9, 3).value == "VeryHigh"
    assert ws.cell(9, 3).font == ws.cell(8, 3).font

    # 既存の TheftRating 行 (6-8) と、後続の Datatype 宣言行は1行下にシフト
    for r in (6, 7, 8):
        assert ws.cell(r, 3).value == before_ws.cell(r, 3).value
    assert ws.cell(12, 3).value == before_ws.cell(11, 3).value  # "Datatype InjuryRating <String>"


def test_patch_write_add_spreadsheet_step(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(SHOP_POLICY)
    table = edited.get_table("Calculation")
    table.steps.append(SpreadsheetStep(label="Test", value="= 1 + 1", unit="テスト"))

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, SHOP_POLICY, out_path)

    wb = openpyxl.load_workbook(out_path)
    ws = wb["Calculation"]

    assert ws.max_row == 6
    assert ws.cell(6, 2).value == "Test"
    assert ws.cell(6, 3).value == "テスト"
    assert ws.cell(6, 4).value == "= 1 + 1"
    assert ws.cell(6, 4).data_type == "s"
    for col in (2, 3, 4):
        assert ws.cell(6, col).font == ws.cell(5, col).font
```

- [ ] **Step 2: テストを実行して失敗を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py -v -k "append_rule or insert_rule or delete_rule or data_table_row or spreadsheet_step"`
Expected: いずれかが FAIL する可能性がある（オフバイワンや `_row_cell_values` の列順序の誤りなど）。失敗内容を確認する。

- [ ] **Step 3: 必要に応じて `openl/patch_writer.py` を修正する**

Step 2 で失敗したテストがあれば、`_patch_table_rows` / `_insert_table_rows` のロジック（特に `anchor` の算出、`item_rows` のインデックス計算、列オフセット）を見直して修正する。

- [ ] **Step 4: テストを実行して成功を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py -v`
Expected: `15 passed`

- [ ] **Step 5: コミット**

```bash
git add openl/patch_writer.py tests/test_patch_writer.py
git commit -m "test: verify row insert/delete with style preservation in patch_write"
```

---

### Task 4: 新規テーブル・新規シートの追記 / テーブル削除

**Files:**
- Modify: `openl/patch_writer.py`
- Modify: `tests/test_patch_writer.py`

`patch_write` をシート単位の処理に書き換える:

- シート内のテーブルを `start_row` の降順で処理する（行の挿入/削除によるズレを避けるため。Task2/3では単一テーブルのシートのみだったため問題化しなかった）
- before側のみに存在するテーブル → 行範囲を削除（`_delete_table`）
- after側のみに存在するテーブル → 既存シート末尾に追記、または新規シートを作成して追記（`_append_table`）。既存シートに追記する場合は空行を挟む

design specのシナリオ6を検証する。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_patch_writer.py` の末尾に追記:

```python
from openl.models import SimpleDecisionTable, ColumnDef


def test_patch_write_new_table_new_sheet_and_deleted_table(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(SHOP_POLICY)

    # (a) 既存シート(FreeShipping)に新規テーブルを追記
    edited.tables.append(SimpleDecisionTable(
        sheet_name="FreeShipping",
        title="",
        method_signature="SimpleRules Boolean IsExpressShipping (String memberType)",
        table_name="IsExpressShipping",
        conditions=[ColumnDef(name="会員種別", col_type="String", role="condition")],
        results=[ColumnDef(name="速達無料", col_type="String", role="result")],
        rules=[Rule(id=1, conditions={"会員種別": "プレミアム会員"}, results={"速達無料": True})],
        start_col=1,
    ))

    # (b) 新規シート(Promotions)に新規テーブルを追記
    edited.tables.append(SimpleDecisionTable(
        sheet_name="Promotions",
        title="",
        method_signature="SimpleRules Boolean IsPromotionEligible (String memberType)",
        table_name="IsPromotionEligible",
        conditions=[ColumnDef(name="会員種別", col_type="String", role="condition")],
        results=[ColumnDef(name="対象", col_type="String", role="result")],
        rules=[Rule(id=1, conditions={"会員種別": "一般会員"}, results={"対象": True})],
        start_col=1,
    ))

    # (c) CampaignTarget のテーブルを削除
    edited.tables = [t for t in edited.tables if t.sheet_name != "CampaignTarget"]

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, SHOP_POLICY, out_path)

    wb = openpyxl.load_workbook(out_path)

    # (a) FreeShipping: 元の11行 + 空行1 + 新規テーブル3行 = 15行
    ws = wb["FreeShipping"]
    assert ws.max_row == 15
    assert ws.cell(12, 2).value is None  # 区切りの空行
    assert ws.cell(13, 2).value == "SimpleRules Boolean IsExpressShipping (String memberType)"
    assert [ws.cell(14, c).value for c in (2, 3)] == ["会員種別", "速達無料"]
    assert [ws.cell(15, c).value for c in (2, 3)] == ["プレミアム会員", True]

    # (b) Promotions: 新規シート、区切り無しで3行
    assert "Promotions" in wb.sheetnames
    ws_promo = wb["Promotions"]
    assert ws_promo.cell(1, 2).value == "SimpleRules Boolean IsPromotionEligible (String memberType)"
    assert [ws_promo.cell(2, c).value for c in (2, 3)] == ["会員種別", "対象"]
    assert [ws_promo.cell(3, c).value for c in (2, 3)] == ["一般会員", True]

    # (c) CampaignTarget: テーブルが削除されている
    result = OpenLReader().read(out_path)
    assert result.get_table("CampaignTarget") is None

    # PointRate / Calculation は完全に不変
    before_wb = openpyxl.load_workbook(SHOP_POLICY)
    before = _all_cell_values(before_wb)
    after = _all_cell_values(wb)
    assert after["PointRate"] == before["PointRate"]
    assert after["Calculation"] == before["Calculation"]
```

- [ ] **Step 2: テストを実行して失敗を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py::test_patch_write_new_table_new_sheet_and_deleted_table -v`
Expected: FAIL（新規/削除テーブルが処理されないため、行数や `Promotions` シートが期待値と一致しない）

- [ ] **Step 3: `patch_write` をシート単位の処理に書き換える**

`openl/patch_writer.py` の `patch_write` 関数を以下に置き換える（ヘルパー2つを新規追加）:

```python
def _delete_table(ws: Worksheet, before: ParsedTable) -> None:
    """before側のみに存在するテーブルの行範囲（シグネチャ行〜区切り空行まで）を削除する。"""
    n = before.end_row - before.start_row + 1
    ws.delete_rows(before.start_row, n)


def _append_table(ws: Worksheet, table: AnyTable, *, separator: bool) -> None:
    """afterのみに存在するテーブルをシート末尾に追記する。既存シートなら空行を挟む。"""
    has_content = any(c.value is not None for row in ws.iter_rows() for c in row)
    if separator and has_content:
        ws.append([])
    _WRITER_MAP[table.table_kind](ws, table)


def patch_write(edited: OpenLWorkbook, original_path: str | Path, out_path: str | Path) -> None:
    """編集後の edited を、original_path の書式・レイアウトを保ったまま out_path に書き出す。"""
    wb = openpyxl.load_workbook(str(original_path))
    before_parsed = read_with_positions(original_path)

    before_by_id = {_table_identity(p.table): p for p in before_parsed}
    after_by_id = {_table_identity(t): t for t in edited.tables}

    original_sheets = set(wb.sheetnames)
    sheet_names = {p.table.sheet_name for p in before_parsed} | \
        {t.sheet_name for t in edited.tables}

    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(title=sheet_name)
        ws = wb[sheet_name]

        # 行の挿入/削除によるズレを避けるため、行番号の大きい方から処理する
        sheet_before = sorted(
            (p for p in before_parsed if p.table.sheet_name == sheet_name),
            key=lambda p: p.start_row,
            reverse=True,
        )
        handled: set[TableIdentity] = set()

        for parsed in sheet_before:
            ident = _table_identity(parsed.table)
            handled.add(ident)
            after = after_by_id.get(ident)
            if after is None:
                _delete_table(ws, parsed)
            else:
                _patch_table_rows(ws, parsed, after)

        for ident, table in after_by_id.items():
            if ident[0] == sheet_name and ident not in handled:
                _append_table(ws, table, separator=sheet_name in original_sheets)

    wb.save(str(out_path))
```

- [ ] **Step 4: テストを実行して成功を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py -v`
Expected: `16 passed`

- [ ] **Step 5: コミット**

```bash
git add openl/patch_writer.py tests/test_patch_writer.py
git commit -m "feat: handle new/deleted tables and sheets in patch_write"
```

---

### Task 5: 列構成変更時の delete+recreate フォールバック

**Files:**
- Modify: `openl/patch_writer.py`
- Modify: `tests/test_patch_writer.py`

`conditions`/`results`/`columns`/`column_names` が before/after で異なる場合、行のinsert/deleteでは表現できないため、そのテーブルの行範囲を削除して `_write_*` 関数で再作成する（このテーブルのみ書式はopenpyxlデフォルトに戻る。他のテーブル・シートは影響を受けない）。design specのシナリオ7を検証する。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_patch_writer.py` の末尾に追記:

```python
def test_patch_write_condition_column_added_recreates_table(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(SHOP_POLICY)
    table = edited.get_table("FreeShipping")
    table.conditions.append(ColumnDef(name="新条件", col_type="String", role="condition"))
    for rule in table.rules:
        rule.conditions["新条件"] = "test"

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, SHOP_POLICY, out_path)

    wb = openpyxl.load_workbook(out_path)
    ws = wb["FreeShipping"]

    # 行数は変わらない（9ルール + シグネチャ行 + ヘッダー行 = 11行）
    assert ws.max_row == 11
    assert [ws.cell(2, c).value for c in (2, 3, 4, 5)] == ["会員種別", "購入金額", "新条件", "送料無料"]
    assert ws.cell(3, 4).value == "test"
    assert ws.cell(11, 4).value == "test"

    # 再パースすると新しい列構成で読み取れる
    result = OpenLReader().read(out_path)
    fs = result.get_table("FreeShipping")
    assert [c.name for c in fs.conditions] == ["会員種別", "購入金額", "新条件"]
    assert len(fs.rules) == 9
    assert fs.rules[0].conditions["新条件"] == "test"

    # 他のシートは完全に不変
    before_wb = openpyxl.load_workbook(SHOP_POLICY)
    before = _all_cell_values(before_wb)
    after = _all_cell_values(wb)
    for sheet in ("PointRate", "CampaignTarget", "Calculation"):
        assert after[sheet] == before[sheet]
```

- [ ] **Step 2: テストを実行して失敗を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py::test_patch_write_condition_column_added_recreates_table -v`
Expected: FAIL（列構成変更が `_patch_table_rows` に渡され、`_comparison_tuples`/`_row_cell_values` の要素数不一致で `IndexError` 等になる）

- [ ] **Step 3: `_recreate_table` を実装し `patch_write` から呼び出す**

`openl/patch_writer.py` に `_recreate_table` を追加する（`_append_table` の直後）:

```python
def _recreate_table(ws: Worksheet, before: ParsedTable, after: AnyTable) -> None:
    """列構成が変わったテーブルを削除して同じ位置に再作成する。

    このテーブルのみ書式が openpyxl デフォルトに戻る（他のテーブル・セルは無影響）。
    """
    n_before = before.end_row - before.start_row + 1
    ws.delete_rows(before.start_row, n_before)

    tmp_ws = openpyxl.Workbook().active
    _WRITER_MAP[after.table_kind](tmp_ws, after)
    n_after = tmp_ws.max_row

    ws.insert_rows(before.start_row, n_after)
    for r in range(1, n_after + 1):
        for c in range(1, tmp_ws.max_column + 1):
            _set_cell(ws, before.start_row + r - 1, c, tmp_ws.cell(row=r, column=c).value)
```

`patch_write` 内のテーブル振り分け部分を以下に変更する（`_delete_table` の分岐に `elif` を追加):

```python
        for parsed in sheet_before:
            ident = _table_identity(parsed.table)
            handled.add(ident)
            after = after_by_id.get(ident)
            if after is None:
                _delete_table(ws, parsed)
            elif _structure_key(parsed.table) != _structure_key(after):
                _recreate_table(ws, parsed, after)
            else:
                _patch_table_rows(ws, parsed, after)
```

- [ ] **Step 4: テストを実行して成功を確認**

Run: `uv run python -m pytest tests/test_patch_writer.py -v`
Expected: `17 passed`

- [ ] **Step 5: コミット**

```bash
git add openl/patch_writer.py tests/test_patch_writer.py
git commit -m "feat: recreate table on column-structure change in patch_write"
```

---

### Task 6: CLI 振り分け（パッチ/フルリビルド）と統合テスト

**Files:**
- Modify: `openl/cli.py:1-21`（import）, `openl/cli.py:65-78`（`cmd_write`）, `openl/cli.py:108-133`（`main`/argparser）
- Create: `tests/test_cli_write.py`
- Modify: `tests/test_patch_writer.py`

`cmd_write` を、元ファイル（`--source` またはJSONの `source_file`）が存在すればパッチモード、存在しなければ既存のフルリビルド（`OpenLWriter`）に振り分ける。design specのシナリオ9（フルリビルドへのフォールバック回帰）と、シナリオ8（`AutoPolicyCalculation.xlsx` を使った統合テスト：Accidents条件 `>2`→`>=10`、タイトル行のテーマカラー・フォント保持、他シート不変）を検証する。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_cli_write.py` を新規作成:

```python
import argparse
import shutil
from pathlib import Path
from unittest.mock import patch

from openl.cli import cmd_write, _resolve_source_path
from openl.models import OpenLWorkbook
from openl.reader import OpenLReader

SHOP_POLICY_XLSX = Path(__file__).parent.parent / "examples" / "ShopPolicy.xlsx"
SHOP_POLICY_JSON = Path(__file__).parent.parent / "examples" / "ShopPolicy.json"


def test_resolve_source_path_uses_source_file_field():
    wb = OpenLWorkbook(source_file="ShopPolicy.xlsx")
    src = Path("/tmp/work/ShopPolicy.json")
    assert _resolve_source_path(src, wb, None) == Path("/tmp/work/ShopPolicy.xlsx")


def test_resolve_source_path_uses_override():
    wb = OpenLWorkbook(source_file="ShopPolicy.xlsx")
    src = Path("/tmp/work/ShopPolicy.json")
    assert _resolve_source_path(src, wb, "/other/dir/Renamed.xlsx") == Path("/other/dir/Renamed.xlsx")


def test_cmd_write_uses_patch_mode_when_source_exists(tmp_path):
    shutil.copy(SHOP_POLICY_XLSX, tmp_path / "ShopPolicy.xlsx")
    shutil.copy(SHOP_POLICY_JSON, tmp_path / "ShopPolicy.json")

    args = argparse.Namespace(input=str(tmp_path / "ShopPolicy.json"), out=None, source=None)

    with patch("openl.cli.patch_write") as mock_patch, patch("openl.cli.OpenLWriter") as mock_writer:
        cmd_write(args)

    mock_patch.assert_called_once()
    mock_writer.return_value.write.assert_not_called()


def test_cmd_write_falls_back_to_full_rebuild_when_source_missing(tmp_path):
    shutil.copy(SHOP_POLICY_JSON, tmp_path / "ShopPolicy.json")  # .xlsx は配置しない

    args = argparse.Namespace(input=str(tmp_path / "ShopPolicy.json"), out=None, source=None)

    with patch("openl.cli.patch_write") as mock_patch:
        cmd_write(args)

    mock_patch.assert_not_called()
    out = tmp_path / "ShopPolicy.xlsx"
    assert out.exists()
    assert len(OpenLReader().read(out).tables) == 4
```

`tests/test_patch_writer.py` の末尾に追記:

```python
def test_patch_write_integration_auto_policy_accidents_condition(tmp_path):
    from openl.patch_writer import patch_write

    edited = OpenLReader().read(AUTO_POLICY)
    # 注意: AutoPolicyCalculation.xlsx には Vocabulary シートにも
    # table_name="DriverRisk" の DataTable (Datatype) が存在するため、
    # table_kind で SimpleDecisionTable に絞り込む。
    table = next(
        t for t in edited.tables
        if t.table_kind == "SimpleDecisionTable" and t.table_name == "DriverRisk"
    )
    rule = next(r for r in table.rules if r.id == 2)
    assert rule.conditions["Accidents"] == ">2"
    rule.conditions["Accidents"] = ">=10"

    out_path = tmp_path / "out.xlsx"
    patch_write(edited, AUTO_POLICY, out_path)

    before_wb = openpyxl.load_workbook(AUTO_POLICY)
    after_wb = openpyxl.load_workbook(out_path)
    before_ws = before_wb["Driver-Eligibility"]
    after_ws = after_wb["Driver-Eligibility"]

    # (a) 値が正しく変わる
    assert before_ws.cell(28, 4).value == ">2"
    assert after_ws.cell(28, 4).value == ">=10"

    # (b) ヘッダー行のBoldフォント・テーマカラー（背景塗りつぶし）が保持される
    before_header = before_ws.cell(26, 4)
    after_header = after_ws.cell(26, 4)
    assert after_header.font.bold is True and before_header.font.bold is True
    assert after_header.fill.fgColor.theme == before_header.fill.fgColor.theme
    assert after_header.fill.fgColor.tint == before_header.fill.fgColor.tint

    # (c) 編集対象外の行数・列幅・行高は不変
    assert after_ws.max_row == before_ws.max_row == 31
    for r in range(25, 32):
        assert after_ws.row_dimensions[r].height == before_ws.row_dimensions[r].height
    for col in "ABCDEFG":
        assert after_ws.column_dimensions[col].width == before_ws.column_dimensions[col].width

    # 他のシートは完全に不変
    assert _all_cell_values(after_wb)["Vehicle-Eligibility"] == _all_cell_values(before_wb)["Vehicle-Eligibility"]
```

- [ ] **Step 2: テストを実行して失敗を確認**

Run: `uv run python -m pytest tests/test_cli_write.py tests/test_patch_writer.py::test_patch_write_integration_auto_policy_accidents_condition -v`
Expected: FAIL（`_resolve_source_path`/`patch_write`参照エラー、または `cmd_write` が常に `OpenLWriter` を呼ぶため `mock_writer.return_value.write` が呼ばれてしまう）

- [ ] **Step 3: `openl/cli.py` を修正する**

`openl/cli.py:18-20` のimportに `patch_write` を追加:

```python
from .models import OpenLWorkbook
from .reader import OpenLReader
from .writer import OpenLWriter
from .patch_writer import patch_write
```

`openl/cli.py:65-78` の `cmd_write` を以下に置き換える:

```python
def _resolve_source_path(src: Path, wb: OpenLWorkbook, source_override: str | None) -> Path:
    """元のExcelファイルパスを解決する。--source指定があればそれを優先する。"""
    if source_override:
        return Path(source_override)
    return src.parent / wb.source_file


def cmd_write(args: argparse.Namespace) -> None:
    """JSON / YAML → Excel（元ファイルが見つかればパッチモード、なければフルリビルド）"""
    src = Path(args.input)
    wb = _load_workbook(src)

    if args.out:
        out = Path(args.out)
    else:
        out = src.with_suffix(".xlsx")

    original_path = _resolve_source_path(src, wb, args.source)
    if original_path.exists():
        patch_write(wb, original_path, out)
        print(f"パッチモード: {original_path} → {out}")
    else:
        OpenLWriter().write(wb, out)
        print(f"元ファイルが見つからないためフルリビルドします（{original_path}）")

    print(f"読み込み完了: {src}")
    print(f"出力        : {out}")
    _print_summary(wb)
```

`openl/cli.py` の `write` サブパーサー定義（`p_write = sub.add_parser(...)` の直後）に `--source` を追加:

```python
    p_write = sub.add_parser("write", help="JSON / YAML → Excel")
    p_write.add_argument("input", help="入力ファイル (.json / .yaml)")
    p_write.add_argument("--out", help="出力Excelパス（省略時: 入力と同ディレクトリ）")
    p_write.add_argument("--source", help="元のExcelファイルパス（省略時: 入力と同ディレクトリの source_file）")
```

- [ ] **Step 4: テストを実行して成功を確認**

Run: `uv run python -m pytest -v`
Expected: 全テストが pass する（`tests/test_patch_writer.py` 18件、`tests/test_cli_write.py` 4件を含む、既存テストの回帰なし）

- [ ] **Step 5: コミット**

```bash
git add openl/cli.py tests/test_cli_write.py tests/test_patch_writer.py
git commit -m "feat: dispatch openl write between patch mode and full rebuild"
```

---

## 完了後の確認

- `uv run python -m openl.cli write examples/ShopPolicy.json --out /tmp/ShopPolicy_patched.xlsx` を実行し、`openl read` で再読込して構造が変わらないことを目視確認する
- `skills/openl-tablets-edit/SKILL.md` のStep 5（書き戻し手順）に変更が必要か確認する（`--source` オプションの説明追加が望ましいが、デフォルト動作のみ使う場合は変更不要）

