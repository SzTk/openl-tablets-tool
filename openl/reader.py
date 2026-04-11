"""
OpenL Tablets Excel Reader

Excel シートを解析し OpenLWorkbook モデルに変換する。

テーブル種別の判定ロジック:
  シートの先頭データ行（Noneをスキップ）を走査し、
  B列 (index 1) の値で判定する。

  'Rules'       → SimpleDecisionTable
  'Data'        → DataTable
  'Spreadsheet' → SpreadsheetTable (フォールバック)
"""

from __future__ import annotations
from pathlib import Path
from typing import Any
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from .models import (
    CellStyle,
    SheetDimensions,
    _DEFAULT_FONT_NAME,
    _DEFAULT_FONT_SIZE,
    ColumnDef,
    Rule,
    SimpleDecisionTable,
    DataTable,
    DataTableRow,
    SpreadsheetTable,
    SpreadsheetParam,
    SpreadsheetStep,
    OpenLWorkbook,
)


def _cell_value(ws: Worksheet, row: int, col: int) -> Any:
    """1始まり行・列でセル値を取得。"""
    return ws.cell(row=row, column=col).value


def _rows_as_lists(ws: Worksheet) -> list[list[Any]]:
    """シート全行を list[list] で返す（1始まりインデックスは使わず 0始まりリスト）。"""
    return [
        [cell.value for cell in row]
        for row in ws.iter_rows()
    ]


def _find_table_header_row(rows: list[list[Any]]) -> int | None:
    """
    'Rules' / 'Data' / 'Spreadsheet' が B列(index=1)にある行インデックスを返す。
    見つからなければ None。
    """
    keywords = {"Rules", "Data", "Spreadsheet"}
    for i, row in enumerate(rows):
        if len(row) > 1 and row[1] in keywords:
            return i
    return None


# ---------------------------------------------------------------------------
# SimpleDecisionTable パーサ
# ---------------------------------------------------------------------------

def _parse_simple_decision_table(sheet_name: str, rows: list[list[Any]]) -> SimpleDecisionTable:
    """
    行レイアウト:
      rows[0] : タイトル行
      rows[1] : Module/Method宣言行
      rows[2] : Rules | SimpleDecisionTable | テーブル名
      rows[3] : Condition ... | Result ... (役割ラベル行)
      rows[4] : ID | 列名(型) ... (ヘッダ行)
      rows[5+]: データ行
    """
    title = str(rows[0][1] or "").strip()
    method_sig = str(rows[1][1] or "").strip()
    table_name = str(rows[2][3] or "").strip()

    # 役割ラベル行 (row index 3)
    role_row = rows[3]
    # ヘッダ行 (row index 4)
    header_row = rows[4]

    # B列以降の有効列数を特定（ヘッダ行でNoneでない列）
    # ID列(index=1)を除いて解析
    conditions: list[ColumnDef] = []
    results: list[ColumnDef] = []

    # 役割ラベルの解析: "Condition" が続く列は condition、"Result" 以降は result
    # role_row: [None, 'Condition', None, None, None, 'Result', '備考']
    # header_row: [None, 'ID', 'Origin (String)', ..., 'BasePrice (double)', '説明']

    current_role = "condition"
    note_col_index: int | None = None

    for col_idx in range(2, len(header_row)):  # IDの次の列から
        role_label = role_row[col_idx] if col_idx < len(role_row) else None
        header = header_row[col_idx] if col_idx < len(header_row) else None

        if role_label == "Result":
            current_role = "result"
        elif role_label in ("備考", "Note", "Notes"):
            note_col_index = col_idx
            break

        if header is None:
            continue

        col_def = ColumnDef.parse(str(header), role=current_role)
        if current_role == "condition":
            conditions.append(col_def)
        else:
            results.append(col_def)

    # データ行の解析 (row index 5 以降)
    rules: list[Rule] = []
    for row in rows[5:]:
        if all(v is None for v in row):
            continue
        rule_id = row[1]
        if rule_id is None:
            continue

        cond_values: dict[str, Any] = {}
        result_values: dict[str, Any] = {}

        # 条件列: index 2 〜 2+len(conditions)-1
        for i, col_def in enumerate(conditions):
            raw = row[2 + i] if (2 + i) < len(row) else None
            cond_values[col_def.name] = raw

        # 結果列: index 2+len(conditions) 〜
        result_start = 2 + len(conditions)
        for i, col_def in enumerate(results):
            raw = row[result_start + i] if (result_start + i) < len(row) else None
            result_values[col_def.name] = raw

        # 備考列
        note_value = None
        if note_col_index is not None and note_col_index < len(row):
            note_value = row[note_col_index]

        rules.append(Rule(
            id=int(rule_id),
            conditions=cond_values,
            results=result_values,
            notes=str(note_value) if note_value is not None else None,
        ))

    return SimpleDecisionTable(
        sheet_name=sheet_name,
        title=title,
        method_signature=method_sig,
        table_name=table_name,
        conditions=conditions,
        results=results,
        rules=rules,
    )


# ---------------------------------------------------------------------------
# DataTable パーサ
# ---------------------------------------------------------------------------

def _parse_data_table(sheet_name: str, rows: list[list[Any]]) -> DataTable:
    """
    行レイアウト:
      rows[0] : タイトル行
      rows[1] : Data | テーブル型 | テーブル名
      rows[2] : 列名(型) ... ヘッダ行
      rows[3+]: データ行
    """
    title = str(rows[0][1] or "").strip()
    table_type = str(rows[1][2] or "").strip()
    table_name = str(rows[1][3] or "").strip()

    # ヘッダ行 (index 2)
    header_row = rows[2]
    columns: list[ColumnDef] = []
    col_indices: list[int] = []

    for col_idx in range(1, len(header_row)):
        header = header_row[col_idx]
        if header is None:
            continue
        col_def = ColumnDef.parse(str(header), role="data")
        columns.append(col_def)
        col_indices.append(col_idx)

    # データ行 (index 3 以降)
    data_rows: list[DataTableRow] = []
    for row in rows[3:]:
        if all(v is None for v in row):
            continue
        row_data: dict[str, Any] = {}
        for col_def, col_idx in zip(columns, col_indices):
            row_data[col_def.name] = row[col_idx] if col_idx < len(row) else None
        data_rows.append(DataTableRow(data=row_data))

    return DataTable(
        sheet_name=sheet_name,
        title=title,
        table_type=table_type,
        table_name=table_name,
        columns=columns,
        rows=data_rows,
    )


# ---------------------------------------------------------------------------
# SpreadsheetTable パーサ
# ---------------------------------------------------------------------------

def _parse_spreadsheet_table(sheet_name: str, rows: list[list[Any]]) -> SpreadsheetTable:
    """
    FinalPriceCalc のような自由形式シートを解析する。
    B列にラベル、C列に値、D列に説明/単位を持つ行を走査する。
    """
    title = str(rows[0][1] or "").strip()
    description: str | None = None
    parameters: list[SpreadsheetParam] = []
    steps: list[SpreadsheetStep] = []
    notes: list[str] = []

    # 解析ステート
    in_params = False
    in_steps = False

    for row in rows[1:]:
        if all(v is None for v in row):
            continue

        b = row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None
        d = row[3] if len(row) > 3 else None

        if b is None:
            continue

        b_str = str(b).strip()

        # セクション切り替え
        if "入力パラメータ" in b_str:
            in_params = True
            in_steps = False
            continue
        if "計算ステップ" in b_str:
            in_params = False
            in_steps = True
            continue
        if b_str.startswith("※") or b_str.startswith("*"):
            notes.append(b_str)
            continue
        if description is None and c is None and d is None and not in_params and not in_steps:
            description = b_str
            continue

        if in_params and c is not None:
            parameters.append(SpreadsheetParam(
                name=b_str,
                value=c,
                description=str(d).strip() if d else None,
            ))
        elif in_steps and b_str:
            steps.append(SpreadsheetStep(
                label=b_str,
                value=c,
                unit=str(d).strip() if d else None,
            ))

    return SpreadsheetTable(
        sheet_name=sheet_name,
        title=title,
        description=description,
        parameters=parameters,
        steps=steps,
        notes=notes,
    )


# ---------------------------------------------------------------------------
# スタイル・サイズ読み取り
# ---------------------------------------------------------------------------

_DEFAULT_FONT_COLORS = {"FF000000", "00000000", ""}   # 黒・透明 = デフォルト扱い
_DEFAULT_FILL_COLORS = {"00000000", "FFFFFFFF", ""}


def _border_style(side) -> str | None:
    """openpyxl の Side オブジェクトからスタイル文字列を返す。なければ None。"""
    return side.style if side and side.style else None


def _read_sheet_styles(ws: Worksheet) -> dict[str, CellStyle]:
    """
    ワークシートの全セルを走査し、デフォルト以外の書式を持つセルのスタイルを返す。
    戻り値: {セル番地 (例: "B1") → CellStyle}
    """
    styles: dict[str, CellStyle] = {}

    for row in ws.iter_rows():
        for cell in row:
            font = cell.font
            fill = cell.fill
            border = cell.border
            nf = cell.number_format or "General"

            # フォント名（デフォルト=Calibri は None として扱う）
            font_name: str | None = None
            if font.name and font.name != _DEFAULT_FONT_NAME:
                font_name = font.name

            # フォントサイズ（デフォルト=11.0pt は None として扱う）
            font_size: float | None = None
            if font.size and font.size != _DEFAULT_FONT_SIZE:
                font_size = float(font.size)

            # フォントカラー
            font_color: str | None = None
            if font.color and font.color.type == "rgb":
                rgb = font.color.rgb
                if rgb not in _DEFAULT_FONT_COLORS:
                    font_color = rgb

            # 塗りつぶし（solid のみ対象）
            fill_color: str | None = None
            if fill.fill_type == "solid" and fill.fgColor.type == "rgb":
                rgb = fill.fgColor.rgb
                if rgb not in _DEFAULT_FILL_COLORS:
                    fill_color = rgb

            # 罫線
            b_top    = _border_style(border.top)
            b_bottom = _border_style(border.bottom)
            b_left   = _border_style(border.left)
            b_right  = _border_style(border.right)

            style = CellStyle(
                bold=bool(font.bold),
                italic=bool(font.italic),
                font_name=font_name,
                font_size=font_size,
                font_color=font_color,
                fill_color=fill_color,
                number_format=nf,
                border_top=b_top,
                border_bottom=b_bottom,
                border_left=b_left,
                border_right=b_right,
            )
            if not style.is_default():
                styles[cell.coordinate] = style

    return styles


def _read_sheet_dimensions(ws: Worksheet) -> SheetDimensions | None:
    """列幅・行高をシートから読み取る。設定がなければ None を返す。"""
    col_widths: dict[str, float] = {}
    for col_letter, col_dim in ws.column_dimensions.items():
        if col_dim.width:
            col_widths[col_letter] = col_dim.width

    row_heights: dict[str, float] = {}
    for row_num, row_dim in ws.row_dimensions.items():
        if row_dim.height:
            row_heights[str(row_num)] = row_dim.height

    if not col_widths and not row_heights:
        return None
    return SheetDimensions(column_widths=col_widths, row_heights=row_heights)


# ---------------------------------------------------------------------------
# メインリーダ
# ---------------------------------------------------------------------------

TABLE_KIND_MAP = {
    "Rules": _parse_simple_decision_table,
    "Data": _parse_data_table,
    "Spreadsheet": _parse_spreadsheet_table,
}


class OpenLReader:
    def read(self, path: str | Path) -> OpenLWorkbook:
        wb = openpyxl.load_workbook(str(path), data_only=True)
        workbook = OpenLWorkbook(source_file=str(Path(path).name))

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = _rows_as_lists(ws)

            # 空シートはスキップ
            if all(all(v is None for v in row) for row in rows):
                continue

            # テーブルデータの解析
            header_row_idx = _find_table_header_row(rows)
            if header_row_idx is None:
                table = _parse_spreadsheet_table(sheet_name, rows)
            else:
                kind_key = rows[header_row_idx][1]
                parser = TABLE_KIND_MAP.get(kind_key, _parse_spreadsheet_table)
                start = max(0, header_row_idx - 2)
                table = parser(sheet_name, rows[start:])

            workbook.tables.append(table)

            # 書式の読み取り
            sheet_style = _read_sheet_styles(ws)
            if sheet_style:
                workbook.sheet_styles[sheet_name] = sheet_style

            # 列幅・行高の読み取り
            dims = _read_sheet_dimensions(ws)
            if dims:
                workbook.sheet_dimensions[sheet_name] = dims

        return workbook
