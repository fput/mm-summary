"""Synthetic sample transactions for docs and screenshots."""

from __future__ import annotations

import datetime as dt

from models import TaxonomyConfig

__all__ = [
    "SAMPLE_TAXONOMY",
    "generate_sample_raw_transactions",
]

SAMPLE_INCOME_GROUP = "Income"
SAMPLE_EXPENSE_GROUP_ORDER = (
    "Housing",
    "Living",
    "Mobility",
    "Technology",
    "Other",
)
SAMPLE_CATEGORY_PRIORITY_BY_GROUP = {
    SAMPLE_INCOME_GROUP: ("Salary",),
    "Housing": ("Rent",),
    "Mobility": ("Public Transport",),
}
SAMPLE_TAXONOMY = TaxonomyConfig(
    income_group=SAMPLE_INCOME_GROUP,
    expense_group_order=SAMPLE_EXPENSE_GROUP_ORDER,
    category_priority_by_group=SAMPLE_CATEGORY_PRIORITY_BY_GROUP,
    group_aliases={SAMPLE_INCOME_GROUP.lower(): SAMPLE_INCOME_GROUP},
    uncategorized_group="Uncategorized",
    uncategorized_category="(uncategorized)",
    no_subcategory="(no subcategory)",
    routine_patterns=("salary", "rent", "electricity", "internet", "public transport"),
)

SALARY_BY_MONTH = (
    2660,
    2660,
    2680,
    2680,
    2700,
    2700,
    2720,
    2720,
    2740,
    2740,
    2760,
    2790,
)
ELECTRICITY_BY_MONTH = (72, 67, 62, 56, 52, 49, 48, 50, 55, 61, 67, 75)
GROCERIES_BY_MONTH = (320, 305, 326, 318, 330, 336, 332, 338, 325, 320, 328, 345)
DINING_OUT_BY_MONTH = (72, 84, 78, 92, 118, 104, 122, 126, 98, 102, 114, 134)
BONUS_BY_MONTH = {12: 550}
FREELANCE_BY_MONTH = {3: 180, 6: 220, 9: 260}
TRAVEL_BY_MONTH = {5: 290, 8: 1150}
HARDWARE_BY_MONTH = {4: 129, 10: 690}
INSURANCE_BY_MONTH = {1: 145, 7: 145}
HEALTH_BY_MONTH = {2: 85, 11: 220}
GIFTS_BY_MONTH = {5: 100, 12: 250}
PUBLIC_TRANSPORT_BY_MONTH = (58, 58, 58, 58, 58, 58, 58, 58, 58, 58, 58, 58)
BICYCLE_BY_MONTH = {3: 24, 5: 36, 7: 62, 9: 18, 11: 128}
DONATIONS_BY_MONTH = {2: 120, 5: 180, 8: 100, 11: 250}


def generate_sample_raw_transactions(year: int, today=None) -> list[dict[str, object]]:
    if today is None:
        today = dt.date.today()
    last_month = today.month if year == today.year else 12
    raw_transactions: list[dict[str, object]] = []

    def add(
        month: int,
        day: int,
        amount: int,
        category: str,
        name: str,
        purpose: str,
        *,
        comment: str = "",
        booking_text: str,
    ) -> None:
        booking_date = dt.date(year, month, day)
        raw_transactions.append(
            {
                "bookingDate": booking_date,
                "valueDate": booking_date,
                "amount": amount,
                "currency": "EUR",
                "category": category,
                "name": name,
                "purpose": purpose,
                "comment": comment,
                "bookingText": booking_text,
                "booked": True,
                "checkmark": False,
                "id": len(raw_transactions) + 1,
            }
        )

    for month in range(1, last_month + 1):
        add(
            month,
            28,
            SALARY_BY_MONTH[month - 1],
            f"{SAMPLE_INCOME_GROUP} > Salary",
            "North Peak Studio",
            "Monthly salary",
            booking_text="Bank transfer",
        )
        if month in BONUS_BY_MONTH:
            add(
                month,
                27,
                BONUS_BY_MONTH[month],
                f"{SAMPLE_INCOME_GROUP} > Bonus",
                "North Peak Studio",
                "Annual bonus",
                comment="Year-end bonus",
                booking_text="Bank transfer",
            )
        if month in FREELANCE_BY_MONTH:
            add(
                month,
                19,
                FREELANCE_BY_MONTH[month],
                f"{SAMPLE_INCOME_GROUP} > Freelance",
                "Lighthouse Design",
                "Freelance design work",
                comment="Small freelance project",
                booking_text="Bank transfer",
            )

        add(
            month,
            3,
            -890,
            "Housing > Rent",
            "Riverside Property",
            "Monthly rent",
            booking_text="Direct debit",
        )
        add(
            month,
            11,
            -ELECTRICITY_BY_MONTH[month - 1],
            "Housing > Electricity",
            "City Utilities",
            "Electricity bill",
            booking_text="Direct debit",
        )
        add(
            month,
            8,
            -GROCERIES_BY_MONTH[month - 1],
            "Living > Groceries",
            "Fresh Market",
            "Groceries",
            booking_text="Card payment",
        )
        add(
            month,
            15,
            -DINING_OUT_BY_MONTH[month - 1],
            "Living > Dining Out",
            "Neighbourhood Cafe",
            "Dining out",
            booking_text="Card payment",
        )
        add(
            month,
            6,
            -34,
            "Living > Fitness",
            "City Gym",
            "Fitness membership",
            booking_text="Direct debit",
        )
        if month in TRAVEL_BY_MONTH:
            add(
                month,
                18,
                -TRAVEL_BY_MONTH[month],
                "Living > Travel",
                "Rail Europe",
                "Travel booking",
                comment="Long weekend trip" if month == 5 else "Summer trip",
                booking_text="Card payment",
            )

        add(
            month,
            5,
            -PUBLIC_TRANSPORT_BY_MONTH[month - 1],
            "Mobility > Public Transport",
            "Transit Pass",
            "Monthly train and bus pass",
            booking_text="Card payment",
        )
        if month in BICYCLE_BY_MONTH:
            add(
                month,
                24,
                -BICYCLE_BY_MONTH[month],
                "Mobility > Bicycle",
                "Bike Workshop",
                "Bicycle maintenance",
                comment="Bike service" if month == 7 else "",
                booking_text="Card payment",
            )

        add(
            month,
            4,
            -38,
            "Technology > Internet",
            "FiberNet",
            "Home internet",
            booking_text="Direct debit",
        )
        add(
            month,
            21,
            -10,
            "Technology > Software",
            "GitHub",
            "Software subscription",
            booking_text="Card payment",
        )
        if month in HARDWARE_BY_MONTH:
            add(
                month,
                23,
                -HARDWARE_BY_MONTH[month],
                "Technology > Hardware",
                "Desk Store" if month == 4 else "Laptop Market",
                "Hardware upgrade",
                comment="Monitor arm" if month == 4 else "New laptop",
                booking_text="Card payment",
            )

        if month in INSURANCE_BY_MONTH:
            add(
                month,
                9,
                -INSURANCE_BY_MONTH[month],
                "Other > Insurance",
                "Coverage Co",
                "Insurance premium",
                booking_text="Direct debit",
            )
        if month in HEALTH_BY_MONTH:
            add(
                month,
                13,
                -HEALTH_BY_MONTH[month],
                "Other > Health",
                "Health Center",
                "Health expense",
                comment="Dental check" if month == 2 else "New glasses",
                booking_text="Card payment",
            )
        if month in GIFTS_BY_MONTH:
            add(
                month,
                16,
                -GIFTS_BY_MONTH[month],
                "Other > Gifts",
                "Bookshop",
                "Gift shopping",
                comment="Birthday gifts" if month == 5 else "Holiday gifts",
                booking_text="Card payment",
            )
        if month in DONATIONS_BY_MONTH:
            add(
                month,
                22,
                -DONATIONS_BY_MONTH[month],
                "Other > Donations",
                "Local Charity",
                "Donation",
                comment="Fundraiser" if month == 5 else "",
                booking_text="Bank transfer",
            )

    return raw_transactions
