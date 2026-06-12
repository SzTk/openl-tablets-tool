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
