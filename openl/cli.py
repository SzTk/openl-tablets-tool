"""
OpenL Tablets CLI

使い方:
  uv run -m openl.cli read  <input.xlsx>  [--out <output.json|yaml>]
  uv run -m openl.cli write <input.json|yaml> [--out <output.xlsx>]
  uv run -m openl.cli roundtrip <input.xlsx> [--out <output.xlsx>]
"""

from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

import yaml

from .models import OpenLWorkbook
from .reader import OpenLReader
from .writer import OpenLWriter
from .patch_writer import patch_write


def _load_workbook(path: Path) -> OpenLWorkbook:
    """JSON / YAML / Excel のいずれかを読み込む。"""
    suffix = path.suffix.lower()
    if suffix in (".json",):
        data = json.loads(path.read_text(encoding="utf-8"))
        return OpenLWorkbook.model_validate(data)
    if suffix in (".yaml", ".yml"):
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
        return OpenLWorkbook.model_validate(data)
    if suffix in (".xlsx", ".xls"):
        return OpenLReader().read(path)
    raise ValueError(f"未対応の拡張子: {suffix}")


def cmd_read(args: argparse.Namespace) -> None:
    """Excel → JSON / YAML"""
    src = Path(args.input)
    wb = OpenLReader().read(src)

    if args.out:
        out = Path(args.out)
    else:
        fmt = args.format or "json"
        out = src.with_suffix(f".{fmt}")

    suffix = out.suffix.lower()
    if suffix == ".json":
        out.write_text(wb.model_dump_json(indent=2), encoding="utf-8")
    elif suffix in (".yaml", ".yml"):
        out.write_text(
            yaml.dump(wb.model_dump(), allow_unicode=True, sort_keys=False, default_flow_style=False),
            encoding="utf-8",
        )
    else:
        print(f"ERROR: 出力形式 '{suffix}' は未対応（.json / .yaml を指定してください）", file=sys.stderr)
        sys.exit(1)

    print(f"読み込み完了: {src}")
    print(f"出力        : {out}")
    _print_summary(wb)


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


def cmd_roundtrip(args: argparse.Namespace) -> None:
    """Excel → モデル → Excel（動作確認用）"""
    src = Path(args.input)
    wb = OpenLReader().read(src)

    if args.out:
        out = Path(args.out)
    else:
        out = src.with_stem(src.stem + "_roundtrip")

    OpenLWriter().write(wb, out)
    print(f"ラウンドトリップ完了: {src} → {out}")
    _print_summary(wb)


def _print_summary(wb: OpenLWorkbook) -> None:
    print(f"\n  ファイル: {wb.source_file}  テーブル数: {len(wb.tables)}")
    for t in wb.tables:
        kind = t.table_kind
        if kind == "SimpleDecisionTable":
            print(f"  [{t.sheet_name}] {kind}  ルール数: {len(t.rules)}")
        elif kind == "DataTable":
            print(f"  [{t.sheet_name}] {kind}  行数: {len(t.rows)}")
        elif kind == "SpreadsheetTable":
            print(f"  [{t.sheet_name}] {kind}  パラメータ: {len(t.parameters)}  ステップ: {len(t.steps)}")


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="openl",
        description="OpenL Tablets Excel 読み書きツール",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # read
    p_read = sub.add_parser("read", help="Excel → JSON / YAML")
    p_read.add_argument("input", help="入力Excelファイル (.xlsx)")
    p_read.add_argument("--out", help="出力ファイルパス（省略時: 入力と同ディレクトリ）")
    p_read.add_argument("--format", choices=["json", "yaml"], default="json", help="出力形式（既定: json）")

    # write
    p_write = sub.add_parser("write", help="JSON / YAML → Excel")
    p_write.add_argument("input", help="入力ファイル (.json / .yaml)")
    p_write.add_argument("--out", help="出力Excelパス（省略時: 入力と同ディレクトリ）")
    p_write.add_argument("--source", help="元のExcelファイルパス（省略時: 入力と同ディレクトリの source_file）")

    # roundtrip
    p_rt = sub.add_parser("roundtrip", help="Excel → モデル → Excel（動作確認）")
    p_rt.add_argument("input", help="入力Excelファイル (.xlsx)")
    p_rt.add_argument("--out", help="出力Excelパス")

    args = parser.parse_args()
    {"read": cmd_read, "write": cmd_write, "roundtrip": cmd_roundtrip}[args.command](args)


if __name__ == "__main__":
    main()
