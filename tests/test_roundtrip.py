from pathlib import Path

from openl.reader import OpenLReader
from openl.writer import OpenLWriter

FIXTURE = Path(__file__).parent.parent / "examples" / "ShopPolicy.xlsx"


def test_read_shop_policy_finds_all_tables():
    wb = OpenLReader().read(FIXTURE)

    assert len(wb.tables) == 4
    sheet_names = {t.sheet_name for t in wb.tables}
    assert sheet_names == {"Calculation", "FreeShipping", "PointRate", "CampaignTarget"}

    free_shipping = wb.get_table("FreeShipping")
    assert free_shipping.table_kind == "SimpleDecisionTable"
    assert len(free_shipping.rules) == 9


def test_roundtrip_preserves_table_structure(tmp_path):
    wb = OpenLReader().read(FIXTURE)
    out_path = tmp_path / "roundtrip.xlsx"
    OpenLWriter().write(wb, out_path)

    wb2 = OpenLReader().read(out_path)

    assert len(wb2.tables) == len(wb.tables)
    for t1, t2 in zip(wb.tables, wb2.tables):
        assert t1.sheet_name == t2.sheet_name
        assert t1.table_kind == t2.table_kind
