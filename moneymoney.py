"""MoneyMoney export and raw transaction parsing."""

from __future__ import annotations

import datetime as dt
import plistlib
import re
import subprocess
from decimal import Decimal, InvalidOperation
from typing import Any

from models import Money, TaxonomyConfig, Transaction
from user_config import DEFAULT_TAXONOMY

__all__ = [
    "applescript_export_transactions",
    "load_raw_transactions_from_plist",
    "is_booked_raw_transaction",
    "parse_transactions",
]

CATEGORY_SPLIT_RE = re.compile(r"\s*(?:>|›|→|/|\\|:)\s*")


def applescript_export_transactions(
    year: int, account: str | None = None, category: str | None = None
) -> bytes:
    start = f"{year}-01-01"
    end = f"{year}-12-31"

    parts = ['tell application "MoneyMoney" to export transactions']
    if account:
        parts.append(f"from account {_applescript_string(account)}")
    if category:
        parts.append(f"from category {_applescript_string(category)}")
    parts.append(f'from date "{start}" to date "{end}" as "plist"')

    proc = subprocess.run(
        ["osascript", "-e", " ".join(parts)],
        capture_output=True,
        check=False,
    )
    if proc.returncode != 0:
        stderr = proc.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(
            f"MoneyMoney export failed: {stderr or 'unknown AppleScript error'}"
        )
    return proc.stdout


def load_raw_transactions_from_plist(plist_bytes: bytes) -> list[dict[str, Any]]:
    data = plistlib.loads(plist_bytes)

    if isinstance(data, dict):
        txs = data.get("transactions", data.get("Transactions"))
    elif isinstance(data, list):
        txs = data
    else:
        txs = None

    if not isinstance(txs, list):
        raise ValueError(
            "Could not find a transaction list in MoneyMoney plist export."
        )

    return [tx for tx in txs if isinstance(tx, dict)]


def is_booked_raw_transaction(raw: dict[str, Any]) -> bool:
    return raw.get("booked") is not False


def parse_transactions(
    raw_transactions: list[dict[str, Any]],
    taxonomy: TaxonomyConfig = DEFAULT_TAXONOMY,
) -> list[Transaction]:
    parsed: list[Transaction] = []

    for raw in raw_transactions:
        booking_date = _booking_date_from_raw(raw)
        if booking_date is None:
            continue

        group, category = _split_category_path(raw.get("category"), taxonomy=taxonomy)
        parsed.append(
            Transaction(
                booking_date=booking_date,
                group=group,
                category=category,
                amount=_parse_decimal(raw.get("amount")),
                name=_normalize_text(raw.get("name")),
                purpose=_normalize_text(raw.get("purpose")),
                comment=_normalize_text(raw.get("comment")),
                booking_text=_normalize_text(raw.get("bookingText")),
            )
        )

    return parsed


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def _applescript_string(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def _parse_decimal(value: Any) -> Money:
    try:
        return Decimal(str(value or "0"))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _booking_date_from_raw(tx: dict[str, Any]) -> dt.date | None:
    for key in ("bookingDate", "valueDate"):
        value = tx.get(key)
        if isinstance(value, dt.datetime):
            return value.date()
        if isinstance(value, dt.date):
            return value
    return None


def _canonicalize_group(raw_group: str, taxonomy: TaxonomyConfig) -> str:
    group = _normalize_text(raw_group)
    key = group.lower().rstrip(".")
    return taxonomy.group_aliases.get(key, group or taxonomy.uncategorized_group)


def _split_category_path(
    raw_category: Any, taxonomy: TaxonomyConfig
) -> tuple[str, str]:
    text = _normalize_text(raw_category)
    if not text or text in {"(Uncategorized)", "(Unkategorisiert)"}:
        return taxonomy.uncategorized_group, taxonomy.uncategorized_category

    parts = [part.strip() for part in CATEGORY_SPLIT_RE.split(text) if part.strip()]
    if not parts:
        return taxonomy.uncategorized_group, taxonomy.uncategorized_category

    group = _canonicalize_group(parts[0], taxonomy=taxonomy)
    category = (
        " / ".join(parts[1:]).strip() if len(parts) > 1 else taxonomy.no_subcategory
    )
    return group, category or taxonomy.uncategorized_category
