from pathlib import Path

from openl.reader import read_with_positions

SHOP_POLICY = Path(__file__).parent.parent / "examples" / "ShopPolicy.xlsx"
AUTO_POLICY = Path(__file__).parent.parent / "examples" / "AutoPolicyCalculation.xlsx"


def test_read_with_positions_covers_all_sheets():
    parsed_tables = read_with_positions(SHOP_POLICY)
    sheet_names = {p.table.sheet_name for p in parsed_tables}
    assert sheet_names == {"Calculation", "FreeShipping", "PointRate", "CampaignTarget"}


def test_simple_decision_table_positions():
    parsed_tables = read_with_positions(SHOP_POLICY)
    free_shipping = next(p for p in parsed_tables if p.table.sheet_name == "FreeShipping")

    assert free_shipping.start_row == 1
    assert free_shipping.header_rows == 2
    assert free_shipping.item_rows == [3, 4, 5, 6, 7, 8, 9, 10, 11]
    assert free_shipping.end_row == 11
    assert len(free_shipping.table.rules) == len(free_shipping.item_rows)


def test_spreadsheet_table_positions():
    parsed_tables = read_with_positions(SHOP_POLICY)
    calculation = next(p for p in parsed_tables if p.table.sheet_name == "Calculation")

    assert calculation.start_row == 1
    assert calculation.header_rows == 2
    assert calculation.item_rows == [3, 4, 5]
    assert calculation.end_row == 5
    assert len(calculation.table.steps) == len(calculation.item_rows)


def test_data_table_positions():
    parsed_tables = read_with_positions(AUTO_POLICY)
    theft_rating = next(
        p for p in parsed_tables
        if p.table.sheet_name == "Vocabulary" and p.table.table_type == "TheftRating"
    )

    assert theft_rating.start_row == 5
    assert theft_rating.header_rows == 1
    assert theft_rating.item_rows == [6, 7, 8]
    assert theft_rating.end_row == 10
    assert len(theft_rating.table.rows) == len(theft_rating.item_rows)
