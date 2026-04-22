"""
OpenL Tablets データモデル定義

テーブル種別:
  - SimpleDecisionTable : 条件 → 結果のルールテーブル (PriceRules, DiscountRules)
  - DataTable           : マスタデータテーブル (Airports, Constants)
  - SpreadsheetTable    : 自由形式の計算シート (FinalPriceCalc)
"""

from __future__ import annotations
from typing import Any, Literal
from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# セル書式モデル
# ---------------------------------------------------------------------------

_DEFAULT_FONT_NAME = "Calibri"
_DEFAULT_FONT_SIZE = 11.0


class CellStyle(BaseModel):
    """
    1 セル分の書式情報。
    color 値はすべて ARGB 8桁 hex 文字列 (例: "FF1F3864")。
    border_* はスタイル文字列 ("thin", "medium", "thick" など)。None = 罫線なし。
    font_name / font_size が None のときはワークブックデフォルト (Calibri / 11pt) として扱う。
    """
    bold: bool = False
    italic: bool = False
    font_name: str | None = None      # None = Calibri (デフォルト)
    font_size: float | None = None    # None = 11.0pt (デフォルト)
    font_color: str | None = None     # ARGB hex。None = 黒
    fill_color: str | None = None     # ARGB hex。None = 塗りなし
    number_format: str = "General"
    border_top: str | None = None
    border_bottom: str | None = None
    border_left: str | None = None
    border_right: str | None = None

    def is_default(self) -> bool:
        """すべてデフォルト値なら True（保存不要と判定する用途）。"""
        return (
            not self.bold
            and not self.italic
            and self.font_name is None
            and self.font_size is None
            and self.font_color is None
            and self.fill_color is None
            and self.number_format == "General"
            and self.border_top is None
            and self.border_bottom is None
            and self.border_left is None
            and self.border_right is None
        )


class SheetDimensions(BaseModel):
    """シートの列幅・行高設定。"""
    column_widths: dict[str, float] = Field(default_factory=dict)  # 列文字 → 幅
    row_heights: dict[str, float] = Field(default_factory=dict)    # 行番号(str) → 高さ


# ---------------------------------------------------------------------------
# 共通パーツ
# ---------------------------------------------------------------------------

class ColumnDef(BaseModel):
    """列定義。'Origin (String)' → name='Origin', col_type='String' に分解する。"""
    name: str
    col_type: str                          # String / int / double など
    role: Literal["condition", "result", "data", "note"] = "data"

    @staticmethod
    def parse(raw: str, role: str = "data") -> "ColumnDef":
        """'Name (Type)' 形式の文字列を ColumnDef に変換する。"""
        raw = str(raw).strip()
        if "(" in raw and raw.endswith(")"):
            name, _, type_part = raw.rpartition("(")
            return ColumnDef(name=name.strip(), col_type=type_part.rstrip(")").strip(), role=role)
        return ColumnDef(name=raw, col_type="String", role=role)


class Rule(BaseModel):
    """SimpleDecisionTable の 1 行分のルール。"""
    id: int
    conditions: dict[str, Any]   # 列名 → 条件式または値
    results: dict[str, Any]      # 列名 → 結果値
    notes: str | None = None


# ---------------------------------------------------------------------------
# テーブル種別モデル
# ---------------------------------------------------------------------------

class SimpleDecisionTable(BaseModel):
    """
    条件列 + 結果列で構成されるルールテーブル。
    PriceRules / DiscountRules シートに対応。
    """
    table_kind: Literal["SimpleDecisionTable"] = "SimpleDecisionTable"
    sheet_name: str
    title: str
    method_signature: str          # Module宣言行のシグネチャ
    table_name: str                # テーブル識別名
    conditions: list[ColumnDef]    # 条件列定義
    results: list[ColumnDef]       # 結果列定義
    rules: list[Rule]
    start_col: int = 1             # テーブル先頭列の 0-based インデックス


class DataTableRow(BaseModel):
    """DataTable の 1 行。列名 → 値 の辞書。"""
    data: dict[str, Any]


class DataTable(BaseModel):
    """
    マスタデータを保持するテーブル。
    Airports / Constants シートに対応。
    """
    table_kind: Literal["DataTable"] = "DataTable"
    sheet_name: str
    title: str
    table_type: str      # 型名 (例: Airports, Constants)
    table_name: str      # インスタンス名 (例: AirportData, ConstantValue)
    columns: list[ColumnDef]
    rows: list[DataTableRow]
    start_col: int = 1   # テーブル先頭列の 0-based インデックス


class SpreadsheetParam(BaseModel):
    """SpreadsheetTable の入力パラメータ 1 件。"""
    name: str
    value: Any
    description: str | None = None


class SpreadsheetStep(BaseModel):
    """SpreadsheetTable の計算ステップ 1 件。"""
    label: str
    value: Any
    unit: str | None = None


class SpreadsheetTable(BaseModel):
    """
    自由形式の計算シミュレーションシート。
    FinalPriceCalc シートに対応。
    """
    table_kind: Literal["SpreadsheetTable"] = "SpreadsheetTable"
    sheet_name: str
    title: str
    description: str | None = None
    parameters: list[SpreadsheetParam] = Field(default_factory=list)
    steps: list[SpreadsheetStep] = Field(default_factory=list)
    column_names: list[str] = Field(default_factory=lambda: ["Step", "Description", "Value"])
    start_col: int = 1   # テーブル先頭列の 0-based インデックス
    notes: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# ワークブック全体
# ---------------------------------------------------------------------------

AnyTable = SimpleDecisionTable | DataTable | SpreadsheetTable


class OpenLWorkbook(BaseModel):
    """Excel ファイル全体を表すルートモデル。"""
    source_file: str
    tables: list[AnyTable] = Field(default_factory=list)
    # シート名 → {セル番地 → CellStyle}。デフォルト書式のセルは含まない。
    sheet_styles: dict[str, dict[str, CellStyle]] = Field(default_factory=dict)
    # シート名 → 列幅・行高
    sheet_dimensions: dict[str, SheetDimensions] = Field(default_factory=dict)

    def get_table(self, sheet_name: str) -> AnyTable | None:
        return next((t for t in self.tables if t.sheet_name == sheet_name), None)
