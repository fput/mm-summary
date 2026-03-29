"""Shared domain models."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from decimal import Decimal

Month = int
Money = Decimal
NoteKey = tuple[str, str, Month]


@dataclass(frozen=True)
class TaxonomyConfig:
    income_group: str
    expense_group_order: tuple[str, ...]
    category_priority_by_group: dict[str, tuple[str, ...]]
    group_aliases: dict[str, str]
    uncategorized_group: str
    uncategorized_category: str
    no_subcategory: str
    routine_patterns: tuple[str, ...]


@dataclass(frozen=True)
class Transaction:
    booking_date: dt.date
    group: str
    category: str
    amount: Money
    name: str
    purpose: str
    comment: str
    booking_text: str


@dataclass
class SummaryData:
    income_categories: dict[str, dict[Month, Money]]
    expense_groups: dict[str, dict[str, dict[Month, Money]]]
    notes: dict[NoteKey, str]
