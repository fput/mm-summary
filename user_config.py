"""User-defined MoneyMoney taxonomy and sorting configuration.

This example is in German.
"""

from models import TaxonomyConfig

DEFAULT_TAXONOMY = TaxonomyConfig(
    # Exact MoneyMoney group name that should be treated as income.
    income_group="Einnahmen",
    # Expense groups are rendered in this order.
    expense_group_order=(
        "Wohnen",
        "Leben",
        "Technik",
        "Sonstiges",
        "Invest",
        "Unkategorisiert",
    ),
    # Prioritize specific categories within a given group.
    category_priority_by_group={
        "Einnahmen": ("Gehalt",),
        "Wohnen": ("Warmmiete",),
    },
    # Normalize alternate source group names to your preferred display name.
    group_aliases={
        "einnahmen": "Einnahmen",
        "misc": "Sonstiges",
        "misc.": "Sonstiges",
        "uncategorized": "Unkategorisiert",
        "unkategorisiert": "Unkategorisiert",
    },
    # Fallback labels used when MoneyMoney data has no usable group/category path.
    uncategorized_group="Unkategorisiert",
    uncategorized_category="(ohne Kategorie)",
    no_subcategory="(ohne Unterkategorie)",
    # Fragments used to suppress noisy notes for routine recurring transactions.
    routine_patterns=(
        "strom",
        "warmmiete",
        "miete",
        "gehalt",
        "salary",
        "gez",
        "rundfunk",
        "internet",
        "telefon",
        "aldi",
        "supermarkt",
    ),
)
