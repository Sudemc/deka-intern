"""Command-line entry point for BOM validation and ERP workbook export."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

from bom_enricher.workflow import process_bom


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Validate a customer BOM and create an ERP-shaped review workbook."
    )
    parser.add_argument("--input", required=True, type=Path, help="Customer BOM .xlsx file")
    parser.add_argument(
        "--provider",
        choices=("sample", "nexar"),
        default="sample",
        help="Component data source (sample catalog works offline; default: sample)",
    )
    parser.add_argument("--output", required=True, type=Path, help="New output .xlsx file")
    parser.add_argument("--catalog", type=Path, help="JSON catalog path for sample provider")
    parser.add_argument("--sheet", help="Input worksheet name (default: first worksheet)")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        report = process_bom(
            input_path=args.input,
            output_path=args.output,
            provider_name=args.provider,
            catalog_path=args.catalog,
            sheet_name=args.sheet or 0,
        )
    except (OSError, ValueError, RuntimeError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 2

    print(f"Processed rows: {len(report.row_results)}")
    print(
        "Checks: "
        + ", ".join(
            f"{name}: {report.counts[name]}"
            for name in ("unchanged", "added", "corrected", "invalid", "unresolved")
        )
    )
    unresolved_codes = {
        item.part_number.casefold()
        for item in report.unresolved_parts
        if item.part_number.strip()
    }
    print(f"Unresolved product-code candidates: {len(unresolved_codes)}")
    print(f"Output: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
