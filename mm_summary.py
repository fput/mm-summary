#!/usr/bin/env python3
"""CLI entrypoint for generating the MoneyMoney year summary workbook."""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

from excel_rendering import create_workbook
from moneymoney import (
    applescript_export_transactions,
    is_booked_raw_transaction,
    load_raw_transactions_from_plist,
    parse_transactions,
)
from sample_data import (
    SAMPLE_TAXONOMY,
    generate_sample_raw_transactions,
)
from user_config import DEFAULT_TAXONOMY


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export MoneyMoney transactions and build an Excel year summary."
    )
    parser.add_argument(
        "--year",
        type=int,
        default=dt.date.today().year,
        help="Year to export (default: current year)",
    )
    parser.add_argument(
        "--account",
        type=str,
        default=None,
        help="Optional MoneyMoney account name/identifier",
    )
    parser.add_argument(
        "--category",
        type=str,
        default=None,
        help="Optional MoneyMoney category filter",
    )
    parser.add_argument("--output", type=Path, default=None, help="Output .xlsx path")
    parser.add_argument(
        "--no-cell-comments",
        action="store_true",
        help="Do not add Excel comments",
    )
    parser.add_argument(
        "--include-transactions-sheet",
        action="store_true",
        help="Include a Transactions sheet with booked raw export rows",
    )
    parser.add_argument(
        "--sample-data",
        action="store_true",
        help="Generate a built-in sample workbook for docs or screenshots",
    )
    parser.add_argument(
        "--no-expense-heatmap",
        action="store_true",
        help="Disable the subtle heatmap on expense month cells",
    )
    return parser.parse_args(argv)


def _default_output_path(year: int, sample_data: bool) -> Path:
    suffix = "-sample" if sample_data else ""
    return Path.cwd() / f"moneymoney-summary{suffix}-{year}.xlsx"


def _load_raw_transactions(
    year: int,
    *,
    account: str | None,
    category: str | None,
    sample_data: bool,
) -> list[dict[str, object]]:
    if sample_data:
        return generate_sample_raw_transactions(year)

    raw_transactions = load_raw_transactions_from_plist(
        applescript_export_transactions(
            year,
            account=account,
            category=category,
        )
    )
    return [tx for tx in raw_transactions if is_booked_raw_transaction(tx)]


def main(argv: list[str]) -> int:
    args = _parse_args(argv)
    output_path = args.output or _default_output_path(
        args.year, sample_data=args.sample_data
    )
    taxonomy = SAMPLE_TAXONOMY if args.sample_data else DEFAULT_TAXONOMY

    try:
        raw_transactions = _load_raw_transactions(
            args.year,
            account=args.account,
            category=args.category,
            sample_data=args.sample_data,
        )
        transactions = parse_transactions(raw_transactions, taxonomy=taxonomy)
        create_workbook(
            raw_transactions=raw_transactions,
            transactions=transactions,
            year=args.year,
            output_path=output_path,
            add_comments=not args.no_cell_comments,
            include_transactions_sheet=args.include_transactions_sheet,
            expense_heatmap=not args.no_expense_heatmap,
            taxonomy=taxonomy,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise sys.exit(main(sys.argv[1:]))
