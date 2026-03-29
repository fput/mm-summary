"""Summary aggregation, notes, and row-ordering rules."""

from __future__ import annotations

import datetime as dt
import re
from collections import defaultdict
from decimal import Decimal
from typing import Iterable

from models import Money, Month, NoteKey, SummaryData, TaxonomyConfig, Transaction
from user_config import DEFAULT_TAXONOMY

__all__ = [
    "average_context",
    "ordered_expense_groups",
    "category_sort_key",
    "build_summary_data",
]

MONTH_NAME_EN = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)

LEGAL_ENTITY_RE = re.compile(r"\b(gmbh|ag|kg|ug|se|ltd|inc|mbh)\b\.?", re.IGNORECASE)


def average_context(year: int, today=None) -> tuple[int, str]:
    if today is None:
        today = dt.date.today()
    if year == today.year:
        return today.month, f"through {MONTH_NAME_EN[today.month - 1]}"
    return 12, "Jan-Dec"


def ordered_expense_groups(
    expense_groups: dict[str, dict],
    expense_group_order: tuple[str, ...] = DEFAULT_TAXONOMY.expense_group_order,
) -> list[str]:
    preferred = list(expense_group_order)
    present_preferred = [group for group in preferred if group in expense_groups]
    remaining = sorted(
        (group for group in expense_groups if group not in preferred),
        key=str.lower,
    )
    return present_preferred + remaining


def category_sort_key(
    group: str,
    category: str,
    category_priority_by_group: dict[
        str, tuple[str, ...]
    ] = DEFAULT_TAXONOMY.category_priority_by_group,
) -> tuple[int, int, str]:
    priorities = category_priority_by_group.get(group, ())
    lowered = category.lower()

    try:
        priority_index = next(
            index
            for index, preferred in enumerate(priorities)
            if lowered == preferred.lower()
        )
    except StopIteration:
        priority_index = 999

    return priority_index, 1 if category.startswith("(") else 0, lowered


def build_summary_data(
    transactions: list[Transaction], taxonomy: TaxonomyConfig = DEFAULT_TAXONOMY
) -> SummaryData:
    income: dict[str, dict[Month, Money]] = defaultdict(
        lambda: defaultdict(lambda: Decimal("0"))
    )
    expenses: dict[str, dict[str, dict[Month, Money]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(lambda: Decimal("0")))
    )
    cell_transactions: dict[NoteKey, list[Transaction]] = defaultdict(list)

    for tx in transactions:
        month = tx.booking_date.month
        if tx.group == taxonomy.income_group:
            income[tx.category][month] += tx.amount
        else:
            expenses[tx.group][tx.category][month] += tx.amount
        cell_transactions[(tx.group, tx.category, month)].append(tx)

    notes: dict[NoteKey, str] = {}
    for key, txs in cell_transactions.items():
        fragments = _dedupe_keep_order(
            fragment
            for tx in txs
            if (
                fragment := _compact_note_fragment(
                    tx,
                    cell_tx_count=len(txs),
                    routine_patterns=taxonomy.routine_patterns,
                )
            )
            is not None
        )
        if fragments:
            text = ", ".join(fragments)
            notes[key] = text[:997].rstrip() + "…" if len(text) > 1000 else text

    return SummaryData(
        income_categories={
            category: dict(months) for category, months in income.items()
        },
        expense_groups={
            group: {category: dict(months) for category, months in categories.items()}
            for group, categories in expenses.items()
        },
        notes=notes,
    )


def _short_text(text: str, max_len: int) -> str:
    cleaned = re.sub(r"\s+", " ", text).strip(" ,;-")
    if len(cleaned) <= max_len:
        return cleaned
    return cleaned[: max_len - 1].rstrip() + "…"


def _abbreviate_counterparty(tx: Transaction) -> str:
    raw = tx.name or tx.booking_text or tx.purpose
    raw = LEGAL_ENTITY_RE.sub("", raw)
    raw = re.sub(r"\s+", " ", raw).strip(" ,;-")
    return _short_text(raw, max_len=14)


def _is_routine_transaction(tx: Transaction, routine_patterns: tuple[str, ...]) -> bool:
    haystack = " ".join(
        [tx.group, tx.category, tx.name, tx.booking_text, tx.purpose]
    ).lower()
    return any(pattern.lower() in haystack for pattern in routine_patterns)


def _dedupe_keep_order(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        key = item.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def _compact_note_fragment(
    tx: Transaction,
    *,
    cell_tx_count: int,
    routine_patterns: tuple[str, ...],
) -> str | None:
    if tx.comment:
        return _short_text(tx.comment, max_len=28)
    if (
        _is_routine_transaction(tx, routine_patterns=routine_patterns)
        or cell_tx_count <= 1
    ):
        return None
    return _abbreviate_counterparty(tx) or None
