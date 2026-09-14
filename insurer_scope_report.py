"""Build an insurer scope report from three Excel workbooks.

Edit the four paths below before running this script. Each input workbook uses
its fourth Excel row as the column header.
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Final, cast

import pandas as pd

# ---- Edit these paths before running the script. ----
AGE_DEBTOR_PATH: Final = Path("/replace/with/AgeDebtorReport.xlsx")
BILLING_PATH: Final = Path("/replace/with/BillingReport.xlsx")
INSURER_INCOME_PATH: Final = Path("/replace/with/InsurerIncomeReport.xlsx")
OUTPUT_PATH: Final = Path("/replace/with/insurer_scope_report.xlsx")

HEADER_ROW: Final = 3

COVER_COL: Final = "Cover No"
VER_COL: Final = "Ver"
COVER_VER_COL: Final = "Cover No/Ver"
INVOICE_COL: Final = "Invoice No"
INCOME_TYPE_COL: Final = "Income Type"
RI_TYPE_COL: Final = "RI Type"
LOOKUP_REFERENCE_COL: Final = "Lookup Reference"
LEAD_INSURER_COL: Final = "Lead Insurer"
PORTION_COL: Final = "Portion"
COINSURANCE_COL: Final = "是否共保"
IS_LEAD_COL: Final = "是否lead"
SCOPE_COL: Final = "Scope"

BILLING_VALUE_COLUMNS: Final = (
    INCOME_TYPE_COL,
    RI_TYPE_COL,
    LOOKUP_REFERENCE_COL,
)
OUTPUT_COLUMNS: Final = (
    COVER_COL,
    VER_COL,
    INVOICE_COL,
    COINSURANCE_COL,
    IS_LEAD_COL,
    *BILLING_VALUE_COLUMNS,
    SCOPE_COL,
)


def _require_columns(
    frame: pd.DataFrame, required: tuple[str, ...], report_name: str
) -> None:
    missing = [column for column in required if column not in frame.columns]
    if missing:
        formatted = ", ".join(repr(column) for column in missing)
        raise ValueError(f"{report_name} is missing required column(s): {formatted}")


def _text(value: object) -> str:
    """Return stripped text, treating Excel blanks as empty strings."""
    if pd.isna(cast(Any, value)):
        return ""
    return str(value).strip()


def _normalize_cover(value: object, *, context: str) -> str:
    cover = re.sub(r"\s+", "", _text(value)).casefold()
    if not cover:
        raise ValueError(f"Malformed Cover No/Ver in {context}: Cover No is blank")
    return cover


def _normalize_ver(value: object, *, context: str) -> str:
    version = _text(value)
    if not version:
        raise ValueError(f"Malformed Cover No/Ver in {context}: Ver is blank")

    try:
        number = Decimal(version)
    except InvalidOperation:
        return re.sub(r"\s+", "", version).casefold()

    if not number.is_finite():
        raise ValueError(
            f"Malformed Cover No/Ver in {context}: invalid Ver {version!r}"
        )
    if number == number.to_integral_value():
        return str(int(number))
    return format(number.normalize(), "f")


def _key_from_parts(cover: object, version: object, *, context: str) -> str:
    return "/".join(
        (
            _normalize_cover(cover, context=context),
            _normalize_ver(version, context=context),
        )
    )


def _key_from_combined(value: object, *, context: str) -> str | None:
    raw_key = _text(value)
    if not raw_key:
        return None

    parts = raw_key.split("/")
    if len(parts) != 2:
        raise ValueError(
            f"Malformed Cover No/Ver in {context}: expected 'Cover No/Ver', "
            f"got {raw_key!r}"
        )
    return _key_from_parts(parts[0], parts[1], context=context)


def _normalize_invoice(value: object, *, insurer_format: bool = False) -> str:
    invoice = re.sub(r"\s+", "", _text(value)).casefold()
    if insurer_format and invoice:
        invoice = f"i{invoice}"
    return invoice


def _is_expected(value: object, expected: str) -> bool:
    return _text(value).casefold() == expected.casefold()


def _read_reports() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    age_debtor = pd.read_excel(AGE_DEBTOR_PATH, header=HEADER_ROW)
    billing = pd.read_excel(BILLING_PATH, header=HEADER_ROW)
    insurer_income = pd.read_excel(INSURER_INCOME_PATH, header=HEADER_ROW)

    _require_columns(
        age_debtor,
        (COVER_COL, VER_COL, INVOICE_COL),
        AGE_DEBTOR_PATH.name,
    )
    _require_columns(
        billing,
        (COVER_VER_COL, *BILLING_VALUE_COLUMNS),
        BILLING_PATH.name,
    )
    _require_columns(
        insurer_income,
        (COVER_VER_COL, INVOICE_COL, LEAD_INSURER_COL, PORTION_COL),
        INSURER_INCOME_PATH.name,
    )
    return age_debtor, billing, insurer_income


def build_report(
    age_debtor: pd.DataFrame,
    billing: pd.DataFrame,
    insurer_income: pd.DataFrame,
) -> pd.DataFrame:
    """Return the final report without modifying any input frame."""
    age_debtor = age_debtor.copy()
    billing = billing.copy()
    insurer_income = insurer_income.copy()

    age_debtor["_key"] = [
        _key_from_parts(cover, version, context=f"AgeDebtorReport row {index + 5}")
        for index, (cover, version) in enumerate(
            zip(age_debtor[COVER_COL], age_debtor[VER_COL], strict=True)
        )
    ]
    billing["_key"] = [
        _key_from_combined(value, context=f"BillingReport row {index + 5}")
        for index, value in enumerate(billing[COVER_VER_COL])
    ]
    insurer_income["_key"] = [
        _key_from_combined(value, context=f"InsurerIncomeReport row {index + 5}")
        for index, value in enumerate(insurer_income[COVER_VER_COL])
    ]

    nonblank_billing = billing[billing["_key"].notna()]
    duplicate_mask = nonblank_billing["_key"].duplicated(keep=False)
    if duplicate_mask.any():
        duplicate_keys = sorted(set(nonblank_billing.loc[duplicate_mask, "_key"]))
        raise ValueError(
            "BillingReport has duplicate nonblank Cover No/Ver key(s): "
            + ", ".join(str(key) for key in duplicate_keys)
        )

    billing_by_key = nonblank_billing.set_index("_key")
    for column in BILLING_VALUE_COLUMNS:
        age_debtor[column] = age_debtor["_key"].map(billing_by_key[column])

    nonblank_insurer = insurer_income[insurer_income["_key"].notna()].copy()
    insurer_counts = nonblank_insurer.groupby("_key", sort=False).size()
    coinsurance_keys = set(insurer_counts[insurer_counts > 1].index)
    age_debtor[COINSURANCE_COL] = age_debtor["_key"].map(
        lambda key: "Y" if key in coinsurance_keys else "N"
    )

    nonblank_insurer["_invoice_key"] = nonblank_insurer[INVOICE_COL].map(
        lambda value: _normalize_invoice(value, insurer_format=True)
    )
    lead_pairs = {
        (row["_key"], row["_invoice_key"])
        for _, row in nonblank_insurer.iterrows()
        if _is_expected(row[LEAD_INSURER_COL], "Lead")
    }

    def lead_flag(row: pd.Series) -> str:
        if row[COINSURANCE_COL] != "Y":
            return ""
        invoice_key = _normalize_invoice(row[INVOICE_COL])
        return "Y" if (row["_key"], invoice_key) in lead_pairs else "N"

    age_debtor[IS_LEAD_COL] = age_debtor.apply(lead_flag, axis=1)

    def scope(row: pd.Series) -> str:
        ri_type = _text(row[RI_TYPE_COL]).casefold()
        has_ri_type = bool(ri_type) and ri_type not in {"na", "n/a"}
        is_commission = _is_expected(row[INCOME_TYPE_COL], "Commission")
        return "N" if has_ri_type or not is_commission else "Y"

    age_debtor[SCOPE_COL] = age_debtor.apply(scope, axis=1)
    return age_debtor.loc[:, list(OUTPUT_COLUMNS)].copy()


def main() -> None:
    age_debtor, billing, insurer_income = _read_reports()
    report = build_report(age_debtor, billing, insurer_income)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    report.to_excel(OUTPUT_PATH, index=False)
    print(f"Wrote {len(report)} rows to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
