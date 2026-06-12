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
