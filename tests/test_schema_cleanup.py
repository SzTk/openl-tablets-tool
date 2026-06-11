import json
from pathlib import Path

from openl.reader import OpenLReader

FIXTURE = Path(__file__).parent.parent / "examples" / "ShopPolicy.xlsx"


def test_workbook_json_has_no_style_or_dimension_keys():
    wb = OpenLReader().read(FIXTURE)
    dumped = json.loads(wb.model_dump_json())

    assert "sheet_styles" not in dumped
    assert "sheet_dimensions" not in dumped
