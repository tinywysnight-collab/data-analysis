"""Refresh the Camunda columns of the master table from Camunda Report and Comment.

Steps (applied in order, each on the result of the previous one):
1. Collection Camunda No.      <- Reconciliation task whose summary is the Inv.No without "I"
2. Leading insurer Camunda No. <- "-" for Y-Leading / N, leading row's collection for Y
3. Billing Camunda No          <- DocGen / eGlobal Billing "rebill:<inv>" task ("@" = whole Cover+Ver)
4. Rebilled Camunda No.        <- Comment "rebill:<inv>" task
5. Leading insurer Camunda No. <- follows the leading row's rebill task for Y rows without rebill
6-9. Status / Requestor / Ticket Owner / Last Update Date of the collection task
"""

from __future__ import annotations

import math
import re
from collections import defaultdict
from pathlib import Path

import pandas as pd

# Master table columns.
COVER = "Cover"
VER = "Ver"
INV_NO = "Inv.No"
CO_INSURER_FLAG = "Co-insurer Flag"
COLLECTION_NO = "Collection Camunda No."
LEADING_NO = "Leading insurer Camunda No."
REBILLED_NO = "Rebilled Camunda No."
BILLING_NO = "Billing Camunda No"
COLLECTION_STATUS = "Collection Camunda Status"
REQUESTOR = "Requestor of Camunda"
TICKET_OWNER = "Ticket Owner"
LAST_UPDATE_DATE = "Last Update Date"

# Camunda Report columns.
TASK_ID = "Task ID"
REQUESTER_NAME = "Requester Name"
PROCESSOR_ASSIGNED = "Processor Assigned"
TASK_TYPE = "Task Type"
TASK_SUMMARY = "Task Summary"
STATUS = "Status"

# Comment columns.
COMMENT_TASK_ID = "TASK_ID"
COMMENTS = "Comments"

FLAG_LEADING = "Y-Leading"
FLAG_FOLLOWER = "Y"
FLAG_NONE = "N"
PLACEHOLDER = "-"

RECONCILIATION = "reconciliation"
BILLING_TASK_TYPES = frozenset({"docgen billing", "eglobal billing"})
SHARE_WITH_COVER_VER = "@"
_REBILL_PATTERN = re.compile(r"rebill\s*:\s*([A-Za-z0-9]+)", re.IGNORECASE)

# Master column <- Camunda Report column, for steps 6-9.
_TASK_DETAIL_COLUMNS = {
    COLLECTION_STATUS: STATUS,
    REQUESTOR: REQUESTER_NAME,
    TICKET_OWNER: PROCESSOR_ASSIGNED,
    LAST_UPDATE_DATE: LAST_UPDATE_DATE,
}

# Placeholder paths; replace with the real locations.
MASTER_PATH = Path("data/input/master.xlsx")
CAMUNDA_REPORT_PATH = Path("data/input/camunda_report.xlsx")
COMMENT_PATH = Path("data/input/comment.xlsx")
OUTPUT_PATH = Path("data/output/master_refreshed.xlsx")


def _text(value: object) -> str:
    if value is None or value is pd.NA or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def _is_blank(value: object) -> bool:
    return _text(value) == ""


def _inv_digits(inv_no: object) -> str:
    """Inv.No without its leading "I", e.g. I01812332 -> 01812332."""
    text = _text(inv_no)
    return text[1:] if text[:1].upper() == "I" else text


def _rebill_inv(text: object) -> str | None:
    """Invoice number referenced by a "rebill:<inv>" text, or None."""
    match = _REBILL_PATTERN.search(_text(text))
    return match.group(1) if match else None


def _cover_ver(df: pd.DataFrame, index: object) -> tuple[str, str]:
    return _text(df.at[index, COVER]), _text(df.at[index, VER])


def _index_by_inv(df: pd.DataFrame) -> dict[str, object]:
    return {_inv_digits(inv): index for index, inv in df[INV_NO].items()}


def _prepare(df: pd.DataFrame) -> pd.DataFrame:
    return df.copy().fillna("").astype(str)


def step1_collection(master: pd.DataFrame, report: pd.DataFrame) -> pd.DataFrame:
    """Fill empty Collection Camunda No. from the matching Reconciliation task."""
    result = _prepare(master)
    task_by_summary: dict[str, str] = {}
    for _, task in report.iterrows():
        if _text(task[TASK_TYPE]).lower() == RECONCILIATION:
            task_by_summary.setdefault(_text(task[TASK_SUMMARY]), _text(task[TASK_ID]))

    for index, row in result.iterrows():
        if not _is_blank(row[COLLECTION_NO]):
            continue
        task_id = task_by_summary.get(_inv_digits(row[INV_NO]))
        if task_id:
            result.at[index, COLLECTION_NO] = task_id
    return result


def step2_leading(master: pd.DataFrame) -> pd.DataFrame:
    """Set Leading insurer Camunda No. to "-" or the leading row's collection task."""
    result = _prepare(master)
    leading_collection: dict[tuple[str, str], str] = {}
    for index, row in result.iterrows():
        if _text(row[CO_INSURER_FLAG]) == FLAG_LEADING:
            leading_collection.setdefault(_cover_ver(result, index), _text(row[COLLECTION_NO]))

    for index, row in result.iterrows():
        flag = _text(row[CO_INSURER_FLAG])
        if flag in (FLAG_LEADING, FLAG_NONE):
            result.at[index, LEADING_NO] = PLACEHOLDER
        elif flag == FLAG_FOLLOWER:
            current = _text(row[LEADING_NO])
            if current not in ("", PLACEHOLDER):
                continue
            collection = leading_collection.get(_cover_ver(result, index), "")
            if collection:
                result.at[index, LEADING_NO] = collection
    return result


def step3_billing(master: pd.DataFrame, report: pd.DataFrame) -> pd.DataFrame:
    """Fill Billing Camunda No from DocGen / eGlobal Billing rebill tasks."""
    result = _prepare(master)
    index_by_inv = _index_by_inv(result)
    rows_by_cover_ver: dict[tuple[str, str], list[object]] = defaultdict(list)
    for index in result.index:
        rows_by_cover_ver[_cover_ver(result, index)].append(index)

    for _, task in report.iterrows():
        if _text(task[TASK_TYPE]).lower() not in BILLING_TASK_TYPES:
            continue
        summary = _text(task[TASK_SUMMARY])
        inv = _rebill_inv(summary)
        if inv is None or inv not in index_by_inv:
            continue
        index = index_by_inv[inv]
        if SHARE_WITH_COVER_VER in summary:
            targets = rows_by_cover_ver[_cover_ver(result, index)]
        else:
            targets = [index]
        for target in targets:
            result.at[target, BILLING_NO] = _text(task[TASK_ID])
    return result


def step4_rebilled(master: pd.DataFrame, comments: pd.DataFrame) -> pd.DataFrame:
    """Fill Rebilled Camunda No. from "rebill:<inv>" comments."""
    result = _prepare(master)
    index_by_inv = _index_by_inv(result)
    for _, comment in comments.iterrows():
        inv = _rebill_inv(comment[COMMENTS])
        if inv is not None and inv in index_by_inv:
            result.at[index_by_inv[inv], REBILLED_NO] = _text(comment[COMMENT_TASK_ID])
    return result


def step5_leading_rebilled(master: pd.DataFrame) -> pd.DataFrame:
    """For Y rows without rebill, follow the rebill task of the row whose collection they point to."""
    result = _prepare(master)
    rebilled_by_collection: dict[str, str] = {}
    for _, row in result.iterrows():
        collection = _text(row[COLLECTION_NO])
        if collection:
            rebilled_by_collection.setdefault(collection, _text(row[REBILLED_NO]))

    for index, row in result.iterrows():
        if _text(row[CO_INSURER_FLAG]) != FLAG_FOLLOWER or not _is_blank(row[REBILLED_NO]):
            continue
        rebilled = rebilled_by_collection.get(_text(row[LEADING_NO]), "")
        if rebilled:
            result.at[index, LEADING_NO] = rebilled
    return result


def step6_to_9_task_details(master: pd.DataFrame, report: pd.DataFrame) -> pd.DataFrame:
    """Copy Status, Requester, Processor and Last Update Date of the collection task."""
    result = _prepare(master)
    for column in _TASK_DETAIL_COLUMNS:
        if column not in result.columns:
            result[column] = ""

    tasks: dict[str, pd.Series] = {}
    for _, task in report.iterrows():
        tasks.setdefault(_text(task[TASK_ID]), task)

    for index, row in result.iterrows():
        collection_task = tasks.get(_text(row[COLLECTION_NO]))
        if collection_task is None:
            continue
        for master_column, report_column in _TASK_DETAIL_COLUMNS.items():
            result.at[index, master_column] = _text(collection_task[report_column])
    return result


def refresh_master(
    master: pd.DataFrame, report: pd.DataFrame, comments: pd.DataFrame
) -> pd.DataFrame:
    """Apply steps 1-9 and return the refreshed master table."""
    result = step1_collection(master, report)
    result = step2_leading(result)
    result = step3_billing(result, report)
    result = step4_rebilled(result, comments)
    result = step5_leading_rebilled(result)
    return step6_to_9_task_details(result, report)


def _read_excel(path: Path) -> pd.DataFrame:
    # Read everything as text so values like Ver "018" keep their leading zeros.
    return pd.read_excel(path, dtype=str, keep_default_na=False)


def run(master_path: Path, report_path: Path, comment_path: Path, output_path: Path) -> None:
    refreshed = refresh_master(
        _read_excel(master_path), _read_excel(report_path), _read_excel(comment_path)
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    refreshed.to_excel(output_path, index=False)


def main() -> None:
    run(MASTER_PATH, CAMUNDA_REPORT_PATH, COMMENT_PATH, OUTPUT_PATH)


if __name__ == "__main__":
    main()
