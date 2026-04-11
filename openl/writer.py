"""
OpenL Tablets Excel Writer

OpenLWorkbook モデルから OpenL Tablets 形式の Excel ファイルを生成する。

各テーブル種別の行レイアウト:

  SimpleDecisionTable:
    行1 : [None, タイトル]
    行2 : [None, method_signature]
    行3 : [None, "Rules", "SimpleDecisionTable", table_name]
    行4 : [None, "Condition"×n列, "Result"×m列, "備考"]  ← 役割ラベル
    行5 : [None, "ID", "列名 (型)"×n+m, "説明"]          ← ヘッダ
    行6+: [None, id, 条件値×n, 結果値×m, notes]           ← データ

  DataTable:
    行1 : [None, タイトル]
    行2 : [None, "Data", table_type, table_name]
    行3 : [None, "列名 (型)"×n]                           ← ヘッダ
    行4+: [None, 値×n]                                    ← データ

  SpreadsheetTable:
    行1 : [None, タイトル]
    行2 : [None, description]
    行3 : [None, "【入力パラメータ】"]
    行4+: [None, name, value, description]  (パラメータ)
    行n : [None, "【計算ステップ（OpenL式）】"]
    行n+: [None, label, value, unit]        (計算ステップ)
    末尾: [None, note]                      (備考行)
"""

from __future__ import annotations
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.worksheet.worksheet import Worksheet

from .models import (
    CellStyle,
    SheetDimensions,
    _DEFAULT_FONT_NAME,
    _DEFAULT_FONT_SIZE,
    OpenLWorkbook,
    SimpleDecisionTable,
    DataTable,
    SpreadsheetTable,
    AnyTable,
)


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------

def _col_header(name: str, col_type: str) -> str:
    return f"{name} ({col_type})"


def _append_row(ws: Worksheet, values: list) -> None:
    """ws に 1 行追記する。"""
    ws.append(values)


# ---------------------------------------------------------------------------
# テーブル種別ごとのライタ
# ---------------------------------------------------------------------------

def _write_simple_decision_table(ws: Worksheet, table: SimpleDecisionTable) -> None:
    n_cond = len(table.conditions)
    n_result = len(table.results)

    # 行1: タイトル
    _append_row(ws, [None, table.title])

    # 行2: Method署名
    _append_row(ws, [None, table.method_signature])

    # 行3: テーブル宣言
    _append_row(ws, [None, "Rules", "SimpleDecisionTable", table.table_name])

    # 行4: 役割ラベル
    # index: 0=None, 1="Condition"(ID列の上), 2〜2+n_cond-1=None(条件列),
    #        2+n_cond="Result", ..., 末尾="備考"
    role_row: list = [None, "Condition"]
    role_row += [None] * n_cond                        # 全条件列分（ID列はConditionラベル自身が担う）
    role_row += ["Result"] + [None] * (n_result - 1)
    role_row += ["備考"]
    _append_row(ws, role_row)

    # 行5: ヘッダ行
    header_row: list = [None, "ID"]
    header_row += [_col_header(c.name, c.col_type) for c in table.conditions]
    header_row += [_col_header(r.name, r.col_type) for r in table.results]
    header_row += ["説明"]
    _append_row(ws, header_row)

    # 行6+: データ行
    for rule in table.rules:
        data_row: list = [None, rule.id]
        data_row += [rule.conditions.get(c.name) for c in table.conditions]
        data_row += [rule.results.get(r.name) for r in table.results]
        data_row += [rule.notes]
        _append_row(ws, data_row)


def _write_data_table(ws: Worksheet, table: DataTable) -> None:
    # 行1: タイトル
    _append_row(ws, [None, table.title])

    # 行2: テーブル宣言
    _append_row(ws, [None, "Data", table.table_type, table.table_name])

    # 行3: ヘッダ行
    header_row: list = [None] + [_col_header(c.name, c.col_type) for c in table.columns]
    _append_row(ws, header_row)

    # 行4+: データ行
    for row in table.rows:
        data_row: list = [None] + [row.data.get(c.name) for c in table.columns]
        _append_row(ws, data_row)


def _write_spreadsheet_table(ws: Worksheet, table: SpreadsheetTable) -> None:
    # 行1: タイトル
    _append_row(ws, [None, table.title])

    # 行2: 説明
    if table.description:
        _append_row(ws, [None, table.description])

    # 入力パラメータセクション
    _append_row(ws, [None, "【入力パラメータ】"])
    for param in table.parameters:
        _append_row(ws, [None, param.name, param.value, param.description])

    # 計算ステップセクション
    _append_row(ws, [None, "【計算ステップ（OpenL式）】"])
    for step in table.steps:
        _append_row(ws, [None, step.label, step.value, step.unit])

    # 備考
    for note in table.notes:
        _append_row(ws, [None, note])


_WRITER_MAP = {
    "SimpleDecisionTable": _write_simple_decision_table,
    "DataTable": _write_data_table,
    "SpreadsheetTable": _write_spreadsheet_table,
}


# ---------------------------------------------------------------------------
# スタイル・サイズ適用
# ---------------------------------------------------------------------------

def _apply_sheet_styles(ws: Worksheet, styles: dict[str, CellStyle]) -> None:
    """sheet_styles に記録された書式をワークシートの各セルに適用する。"""
    for coord, style in styles.items():
        cell = ws[coord]

        # フォント（name / size / bold / italic / color をまとめて設定）
        # font_color が None の場合は color 引数を渡さず Excel デフォルト（黒）のまま
        font_kwargs: dict = dict(
            name=style.font_name or _DEFAULT_FONT_NAME,
            size=style.font_size or _DEFAULT_FONT_SIZE,
            bold=style.bold,
            italic=style.italic,
        )
        if style.font_color:
            font_kwargs["color"] = style.font_color
        cell.font = Font(**font_kwargs)

        # 塗りつぶし
        if style.fill_color:
            cell.fill = PatternFill(fill_type="solid", fgColor=style.fill_color)

        # 数値書式
        if style.number_format and style.number_format != "General":
            cell.number_format = style.number_format

        # 罫線（4辺を個別に設定）
        if any([style.border_top, style.border_bottom, style.border_left, style.border_right]):
            cell.border = Border(
                top=Side(style=style.border_top) if style.border_top else Side(),
                bottom=Side(style=style.border_bottom) if style.border_bottom else Side(),
                left=Side(style=style.border_left) if style.border_left else Side(),
                right=Side(style=style.border_right) if style.border_right else Side(),
            )


def _apply_sheet_dimensions(ws: Worksheet, dims: SheetDimensions) -> None:
    """列幅・行高をワークシートに適用する。"""
    for col_letter, width in dims.column_widths.items():
        ws.column_dimensions[col_letter].width = width
    for row_num_str, height in dims.row_heights.items():
        ws.row_dimensions[int(row_num_str)].height = height


# ---------------------------------------------------------------------------
# メインライタ
# ---------------------------------------------------------------------------

class OpenLWriter:
    def write(self, workbook: OpenLWorkbook, path: str | Path) -> None:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # デフォルトシートを削除

        for table in workbook.tables:
            ws = wb.create_sheet(title=table.sheet_name)
            writer_fn = _WRITER_MAP.get(table.table_kind)
            if writer_fn is None:
                raise ValueError(f"未対応のテーブル種別: {table.table_kind}")
            writer_fn(ws, table)

            # 書式を適用
            sheet_style = workbook.sheet_styles.get(table.sheet_name, {})
            if sheet_style:
                _apply_sheet_styles(ws, sheet_style)

            # 列幅・行高を適用
            dims = workbook.sheet_dimensions.get(table.sheet_name)
            if dims:
                _apply_sheet_dimensions(ws, dims)

        wb.save(str(path))
