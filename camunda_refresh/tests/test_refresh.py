from pathlib import Path

import pandas as pd
import pytest

from camunda_refresh.refresh import (
    BILLING_NO,
    COLLECTION_NO,
    COLLECTION_STATUS,
    INV_NO,
    LAST_UPDATE_DATE,
    LEADING_NO,
    REBILLED_NO,
    REQUESTOR,
    TICKET_OWNER,
    refresh_master,
    run,
    step1_collection,
    step2_leading,
    step3_billing,
    step4_rebilled,
    step5_leading_rebilled,
    step6_to_9_task_details,
)

MASTER_COLUMNS = [
    "Cover",
    "Ver",
    "Inv.No",
    "Co-insurer Flag",
    "Collection Camunda No.",
    "Leading insurer Camunda No.",
    "Rebilled Camunda No.",
    "Billing Camunda No",
]
MASTER_ROWS = [
    ["1078369", "018", "I01812332", "Y-Leading", "A-AAA0-167734", "N.A", "", ""],
    ["1078369", "018", "I01812333", "Y", "", "A-AAA0-167734", "", ""],
    ["1078372", "017", "I01812340", "Y-Leading", "", "", "", ""],
    ["1078372", "017", "I01812341", "Y", "", "", "", ""],
    ["1930997", "001", "I01809773", "N", "", "", "", ""],
    ["1930997", "000", "I01809770", "N", "", "", "", ""],
    ["1872974", "001", "I01811517", "N", "", "", "", ""],
    ["1930997", "001", "I01809775", "N", "", "", "", ""],
    ["1078369", "020", "I01812336", "Y-Leading", "", "", "", ""],
    ["1078369", "020", "I01812337", "Y", "", "", "", ""],
    ["1078372", "019", "I01812344", "Y-Leading", "", "", "", ""],
]

REPORT_COLUMNS = [
    "Task ID",
    "Created From",
    "Requester Name",
    "Processor Assigned",
    "Task Type",
    "Task Summary",
    "Task Description",
    "Status",
    "Last Update Date",
]
REPORT_ROWS = [
    ["A-AAA0-167734", "-", "Yu, Skye", "Xu, Cissy (U1209125)", "Reconciliation",
     "01812332", "-", "Completed", "2026-09-01"],
    ["A-AAA9-328102", "-", "Yu, Nick", "Yang, Jessie", "Reconciliation",
     "01812340", "-", "Completed", "2026-09-02"],
    ["A-AAAE-119844", "-", "Wu, Eason", "Li, Lillian", "Reconciliation",
     "01812341", "-", "Cancelled", "2026-09-03"],
    ["A-AAAJ-724271", "-", "Dong, Na", "Wu, Ella", "Reconciliation",
     "01809773", "-", "Completed", "2026-09-04"],
    ["A-AAAQ-936846", "-", "iMap_DocGen@marsh.com", "Li, Lillian", "Reconciliation",
     "01809770", "-", "Cancelled", "2026-09-05"],
    ["A-AAAS-016163", "-", "Xiao, Iris", "Xiao, Iris", "DocGen Billing",
     "rebill:01809773", "-", "Completed", "2026-09-06"],
    ["A-AAAX-445913", "-", "iMap_DocGen@marsh.com", "Unassigned", "DocGen Billing",
     "01809775", "-", "Cancelled", "2026-09-07"],
    ["A-AAB0-675253", "-", "Qin, Xiaomin", "Zhang, Xinyu", "eGlobal Billing",
     "rebill:01812332@", "审核无误，待录入", "Completed", "2026-09-08"],
    ["A-JLCY-584598", "-", "Qin, Xiaomin", "Zhang, Xinyu", "Reconciliation",
     "01812336", "", "Inprogress", "2026-09-09"],
    ["A-AAB2-714594", "-", "iMap_DocGen@marsh.com", "Lu, Jennifer", "Reconciliation",
     "01812337", "-", "Cancelled", "2026-09-10"],
    ["A-KDQR-448644", "", "", "", "Reconciliation", "01812344", "", "", ""],
]

COMMENT_COLUMNS = ["TASK_ID", "Comments", "TimeStamp", "User", "Comment Time"]
COMMENT_ROWS = [
    ["A-JLCY-584598", "rebill:01812332", "", "", ""],
    ["A-AAB2-714594", "rebill:01812333", "", "", ""],
    ["A-AAAX-445913", "rebill:01809773", "", "", ""],
    ["A-KDQR-448644", "rebill:01812340", "", "", ""],
]

# Inv.No -> (Collection, Leading, Rebilled, Billing) after the full refresh.
EXPECTED = {
    "I01812332": ("A-AAA0-167734", "-", "A-JLCY-584598", "A-AAB0-675253"),
    "I01812333": ("", "A-AAA0-167734", "A-AAB2-714594", "A-AAB0-675253"),
    "I01812340": ("A-AAA9-328102", "-", "A-KDQR-448644", ""),
    "I01812341": ("A-AAAE-119844", "A-KDQR-448644", "", ""),
    "I01809773": ("A-AAAJ-724271", "-", "A-AAAX-445913", "A-AAAS-016163"),
    "I01809770": ("A-AAAQ-936846", "-", "", ""),
    "I01811517": ("", "-", "", ""),
    "I01809775": ("", "-", "", ""),
    "I01812336": ("A-JLCY-584598", "-", "", ""),
    "I01812337": ("A-AAB2-714594", "A-JLCY-584598", "", ""),
    "I01812344": ("A-KDQR-448644", "-", "", ""),
}


@pytest.fixture(scope="session")
def master() -> pd.DataFrame:
    return pd.DataFrame(MASTER_ROWS, columns=MASTER_COLUMNS)


@pytest.fixture(scope="session")
def report() -> pd.DataFrame:
    return pd.DataFrame(REPORT_ROWS, columns=REPORT_COLUMNS)


@pytest.fixture(scope="session")
def comments() -> pd.DataFrame:
    return pd.DataFrame(COMMENT_ROWS, columns=COMMENT_COLUMNS)


def _value(df: pd.DataFrame, inv_no: str, column: str) -> str:
    return str(df.loc[df[INV_NO] == inv_no, column].iloc[0])


@pytest.mark.parametrize(
    ("inv_no", "expected"),
    [
        ("I01812332", "A-AAA0-167734"),  # already filled, kept
        ("I01812340", "A-AAA9-328102"),
        ("I01809770", "A-AAAQ-936846"),  # cancelled task still counts
        ("I01812333", ""),  # no Reconciliation task
        ("I01809775", ""),  # only a DocGen Billing task matches
    ],
)
def test_step1_fills_empty_collection_from_reconciliation(
    master: pd.DataFrame, report: pd.DataFrame, inv_no: str, expected: str
) -> None:
    result = step1_collection(master, report)
    assert _value(result, inv_no, COLLECTION_NO) == expected


def test_step1_keeps_existing_collection(report: pd.DataFrame) -> None:
    master = pd.DataFrame(
        [["1", "001", "I01812340", "N", "A-KEEP-000001", "", "", ""]], columns=MASTER_COLUMNS
    )
    result = step1_collection(master, report)
    assert _value(result, "I01812340", COLLECTION_NO) == "A-KEEP-000001"


def test_step1_does_not_mutate_input(master: pd.DataFrame, report: pd.DataFrame) -> None:
    before = master.copy()
    step1_collection(master, report)
    pd.testing.assert_frame_equal(master, before)


@pytest.mark.parametrize(
    ("inv_no", "expected"),
    [
        ("I01812332", "-"),  # Y-Leading, N.A replaced
        ("I01809773", "-"),  # N
        ("I01812333", "A-AAA0-167734"),  # Y with existing value, kept
        ("I01812341", "A-AAA9-328102"),  # Y, taken from leading row's collection
        ("I01812337", "A-JLCY-584598"),
    ],
)
def test_step2_leading(
    master: pd.DataFrame, report: pd.DataFrame, inv_no: str, expected: str
) -> None:
    result = step2_leading(step1_collection(master, report))
    assert _value(result, inv_no, LEADING_NO) == expected


def test_step2_refills_follower_marked_dash() -> None:
    master = pd.DataFrame(
        [
            ["1", "001", "I1", "Y-Leading", "A-LEAD", "", "", ""],
            ["1", "001", "I2", "Y", "", "-", "", ""],
        ],
        columns=MASTER_COLUMNS,
    )
    result = step2_leading(master)
    assert _value(result, "I2", LEADING_NO) == "A-LEAD"


@pytest.mark.parametrize(
    ("inv_no", "expected"),
    [
        ("I01809773", "A-AAAS-016163"),  # rebill without @: only this row
        ("I01812332", "A-AAB0-675253"),  # rebill with @
        ("I01812333", "A-AAB0-675253"),  # same Cover + Ver as the @ row
        ("I01812336", ""),  # same Cover, different Ver
        ("I01809775", ""),  # DocGen Billing without rebill
    ],
)
def test_step3_billing(
    master: pd.DataFrame, report: pd.DataFrame, inv_no: str, expected: str
) -> None:
    result = step3_billing(master, report)
    assert _value(result, inv_no, BILLING_NO) == expected


def test_step3_ignores_rebill_with_other_task_type() -> None:
    master = pd.DataFrame([["1", "001", "I1", "N", "", "", "", ""]], columns=MASTER_COLUMNS)
    report = pd.DataFrame(
        [["A-X", "", "", "", "Reconciliation", "rebill:1", "", "", ""]], columns=REPORT_COLUMNS
    )
    assert _value(step3_billing(master, report), "I1", BILLING_NO) == ""


@pytest.mark.parametrize(
    ("inv_no", "expected"),
    [
        ("I01812332", "A-JLCY-584598"),
        ("I01812333", "A-AAB2-714594"),
        ("I01809773", "A-AAAX-445913"),
        ("I01812340", "A-KDQR-448644"),
        ("I01812341", ""),
    ],
)
def test_step4_rebilled(
    master: pd.DataFrame, comments: pd.DataFrame, inv_no: str, expected: str
) -> None:
    result = step4_rebilled(master, comments)
    assert _value(result, inv_no, REBILLED_NO) == expected


def test_step4_ignores_comment_without_rebill() -> None:
    master = pd.DataFrame([["1", "001", "I1", "N", "", "", "", ""]], columns=MASTER_COLUMNS)
    comments = pd.DataFrame([["A-X", "checked 1", "", "", ""]], columns=COMMENT_COLUMNS)
    assert _value(step4_rebilled(master, comments), "I1", REBILLED_NO) == ""


@pytest.mark.parametrize(
    ("row_a", "expected"),
    [
        # Y, no rebill, source row has rebill -> replaced
        (["1", "001", "I2", "Y", "", "A-LEAD", "", ""], "A-NEW"),
        # Y, has rebill -> unchanged
        (["1", "001", "I2", "Y", "", "A-LEAD", "A-OWN", ""], "A-LEAD"),
        # not Y -> unchanged
        (["1", "001", "I2", "N", "", "A-LEAD", "", ""], "A-LEAD"),
        # source row not found -> unchanged
        (["1", "001", "I2", "Y", "", "A-MISSING", "", ""], "A-MISSING"),
    ],
)
def test_step5_leading_follows_rebill(row_a: list[str], expected: str) -> None:
    master = pd.DataFrame(
        [["1", "001", "I1", "Y-Leading", "A-LEAD", "-", "A-NEW", ""], row_a],
        columns=MASTER_COLUMNS,
    )
    assert _value(step5_leading_rebilled(master), "I2", LEADING_NO) == expected


def test_step5_keeps_value_when_source_row_has_no_rebill() -> None:
    master = pd.DataFrame(
        [
            ["1", "001", "I1", "Y-Leading", "A-LEAD", "-", "", ""],
            ["1", "001", "I2", "Y", "", "A-LEAD", "", ""],
        ],
        columns=MASTER_COLUMNS,
    )
    assert _value(step5_leading_rebilled(master), "I2", LEADING_NO) == "A-LEAD"


@pytest.mark.parametrize(
    ("inv_no", "status", "requestor", "owner", "updated"),
    [
        ("I01812332", "Completed", "Yu, Skye", "Xu, Cissy (U1209125)", "2026-09-01"),
        ("I01812336", "Inprogress", "Qin, Xiaomin", "Zhang, Xinyu", "2026-09-09"),
        ("I01812333", "", "", "", ""),  # no collection task
    ],
)
def test_steps6_to_9_task_details(
    master: pd.DataFrame,
    report: pd.DataFrame,
    inv_no: str,
    status: str,
    requestor: str,
    owner: str,
    updated: str,
) -> None:
    result = step6_to_9_task_details(step1_collection(master, report), report)
    assert _value(result, inv_no, COLLECTION_STATUS) == status
    assert _value(result, inv_no, REQUESTOR) == requestor
    assert _value(result, inv_no, TICKET_OWNER) == owner
    assert _value(result, inv_no, LAST_UPDATE_DATE) == updated


def test_refresh_master_matches_target(
    master: pd.DataFrame, report: pd.DataFrame, comments: pd.DataFrame
) -> None:
    result = refresh_master(master, report, comments)
    actual = {
        str(row[INV_NO]): (
            str(row[COLLECTION_NO]),
            str(row[LEADING_NO]),
            str(row[REBILLED_NO]),
            str(row[BILLING_NO]),
        )
        for _, row in result.iterrows()
    }
    assert actual == EXPECTED


def test_run_reads_and_writes_excel_keeping_leading_zeros(
    tmp_path: Path, master: pd.DataFrame, report: pd.DataFrame, comments: pd.DataFrame
) -> None:
    master_path = tmp_path / "master.xlsx"
    report_path = tmp_path / "camunda_report.xlsx"
    comment_path = tmp_path / "comment.xlsx"
    output_path = tmp_path / "out" / "master_refreshed.xlsx"
    master.to_excel(master_path, index=False)
    report.to_excel(report_path, index=False)
    comments.to_excel(comment_path, index=False)

    run(master_path, report_path, comment_path, output_path)

    written = pd.read_excel(output_path, dtype=str, keep_default_na=False)
    assert _value(written, "I01812332", "Ver") == "018"
    assert _value(written, "I01812341", LEADING_NO) == "A-KDQR-448644"
    assert list(written.columns[: len(MASTER_COLUMNS)]) == MASTER_COLUMNS
