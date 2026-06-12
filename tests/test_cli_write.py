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
