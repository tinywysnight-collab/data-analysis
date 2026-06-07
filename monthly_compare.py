"""相邻月份月结报告比对工具（v2）.

自动扫描目录下所有 `YYYYMM.xlsx` 月度月结报告，按时间排序逐对相邻比对，计算
Collection / Working Month / Income Type 三列，并整合 Age_Debt_Report 中“当月即收款、
财务报告里从未出现”的发票，最终把全部结果纵向堆叠输出到单个 Excel 的单个 sheet。

预处理（每个月度文件，按顺序）：
  1. 去无用列：保留第 1–20、35、38 列（按位置，1-indexed），列顺序不变。
  2. 去重：按 Invoice No 分组，Commission AR / Paid / Partial Paid 求和，其余列取第一条，列顺序不变。
  3. 合并 Cover Ver = Cover + "/" + Ver(3位补0)，插在 Ver 列后（保留原 Cover/Ver）；Age_Debt 同样处理。

比对规则（M1=前月，M2=后月，Working Month = M2 年月的 1 号）：
  ① 发票在 M1 有、M2 无：  Total Due(M1) × Rate(M1)
  ② 发票在 M1、M2 都有：    Total Due(M1) − Total Due(M2) × Rate(M1)        （先乘后减）
  ③ 发票在 M2 有、M1 无：    (Paid(M2) + Partial Paid(M2)) × Rate(M2)
  ④ Age_Debt：Invoice Date 属于 M2 当月、且 Invoice No 不在 M2 文件：
                            (Commission AR + Commission VAT) × Rate

Income Type（按每行来源行的 Risk，去首尾空格后不区分大小写）：
  前 3 字母为 Fee 或 前 9 字母为 Multiline → Fee/NPT Fee，否则 → Commission

Report（来源标识）：①② 取 M1 文件名（去扩展名），③ 取 M2 文件名，④ 为 ADR。

Refined Base Equiv：①②④ = Base Equiv（来源行），③ = (Commission AR + Commission VAT) × Rate。

存 Excel 前追加 4 列（按正负拆分）：
  AR = Refined Base Equiv 的正数部分，AP = 其负数部分的相反数；
  Received = Collection 的正数部分，Paid Amount = 其负数部分的相反数。

输出每行 = 保留的原始列(含 Cover/Ver/Cover Ver) … + Collection + Working Month
  + Income Type + Report + Refined Base Equiv + AR + AP + Received + Paid Amount。
"""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

# ---- 列名（按名称引用，需落在保留的列里）----
INVOICE_COL = "Invoice No"
AR_COL = "Commission AR"
VAT_COL = "Commission VAT"
PAID_COL = "Paid"
PARTIAL_COL = "Partial Paid"
RATE_COL = "Currency Rate"
DUE_COL = "Total Due"
RISK_COL = "Risk"
INVOICE_DATE_COL = "Invoice Date"
BASE_EQUIV_COL = "Base Equiv"
COVER_COL = "Cover"
VER_COL = "Ver"

# ---- 新增列 ----
COVER_VER_COL = "Cover Ver"
COLLECTION_COL = "Collection"
WORKING_MONTH_COL = "Working Month"
INCOME_TYPE_COL = "Income Type"
REPORT_COL = "Report"
REFINED_BASE_EQUIV_COL = "Refined Base Equiv"
AR_OUT_COL = "AR"
AP_OUT_COL = "AP"
RECEIVED_COL = "Received"
PAID_AMOUNT_COL = "Paid Amount"

ADR_REPORT = "ADR"  # ④ Age_Debt 行的 Report 取值

# 保留列位置（1-indexed: 1–20, 35, 38）转为 0-indexed
KEEP_POSITIONS = [*range(20), 34, 37]
# 去重时求和的列；其余列取第一条
SUM_COLS = [AR_COL, PAID_COL, PARTIAL_COL]

AGE_DEBT_FILENAME = "Age_Debt_Report.xlsx"
AGE_DEBT_HEADER_ROW = 3  # 0-indexed：前 3 行跳过，第 4 行为表头


def select_kept_columns(df: pd.DataFrame) -> pd.DataFrame:
    """按位置保留第 1–20、35、38 列，保持原列顺序。"""
    return df.iloc[:, KEEP_POSITIONS].copy()


def dedup_invoices(df: pd.DataFrame) -> pd.DataFrame:
    """按 Invoice No 去重：SUM_COLS 求和，其余列取第一条；列顺序不变。"""
    agg = {
        col: ("sum" if col in SUM_COLS else "first")
        for col in df.columns
        if col != INVOICE_COL
    }
    grouped = df.groupby(INVOICE_COL, as_index=False, sort=False).agg(agg)
    return grouped[df.columns.tolist()]  # 恢复原始列顺序


def _format_ver(v: object) -> str:
    """Ver 转 3 位、前面补 0 的字符串（空值按 000，取整数部分）。"""
    if pd.isna(v):
        return "000"
    return str(int(v)).zfill(3)


def add_cover_ver(df: pd.DataFrame) -> pd.DataFrame:
    """新增 Cover Ver = Cover + "/" + Ver(3位补0)，插在 Ver 列后，保留原 Cover/Ver。"""
    cover_ver = df[COVER_COL].astype(str) + "/" + df[VER_COL].map(_format_ver)
    df = df.copy()
    df.insert(df.columns.get_loc(VER_COL) + 1, COVER_VER_COL, cover_ver)
    return df


def read_month(path: Path) -> pd.DataFrame:
    """读取一个月度文件：去无用列 -> 去重 -> 合并 Cover Ver。"""
    df = pd.read_excel(path, dtype={INVOICE_COL: str})
    return add_cover_ver(dedup_invoices(select_kept_columns(df)))


def read_age_debt(path: Path) -> pd.DataFrame:
    """读取 Age_Debt_Report（第 4 行为表头），解析 Invoice Date、合并 Cover Ver。"""
    df = pd.read_excel(path, header=AGE_DEBT_HEADER_ROW, dtype={INVOICE_COL: str})
    df[INVOICE_DATE_COL] = pd.to_datetime(df[INVOICE_DATE_COL], errors="coerce")
    return add_cover_ver(df)


def income_type(risk: object) -> str:
    """Risk 去首尾空格、不区分大小写：Fee.. / Multiline.. -> Fee/NPT Fee，否则 Commission。"""
    s = str(risk).strip().lower()
    if s.startswith("fee") or s.startswith("multiline"):
        return "Fee/NPT Fee"
    return "Commission"


def _make_row(
    src: pd.Series, collection: float, wm: date, report: str, refined: float
) -> dict[str, object]:
    """把来源行 + Collection / Working Month / Income Type / Report / Refined Base Equiv 组成一条记录。"""
    return {
        **src.to_dict(),
        COLLECTION_COL: collection,
        WORKING_MONTH_COL: wm,
        INCOME_TYPE_COL: income_type(src[RISK_COL]),
        REPORT_COL: report,
        REFINED_BASE_EQUIV_COL: refined,
    }


def compare_pair(
    m1: pd.DataFrame, m2: pd.DataFrame, wm: date, m1_name: str, m2_name: str
) -> list[dict[str, object]]:
    """比对相邻两月，返回 ①②③ 的输出记录列表（Report：①②取 M1 名，③取 M2 名）。"""
    m2_by_invoice = m2.set_index(INVOICE_COL)
    m2_invoices = set(m2_by_invoice.index)
    m1_invoices = set(m1[INVOICE_COL])

    rows: list[dict[str, object]] = []

    # ① / ②：遍历 M1
    for _, r in m1.iterrows():
        invoice = r[INVOICE_COL]
        rate1 = r[RATE_COL]
        due1 = r[DUE_COL]
        if invoice in m2_invoices:  # ② 仍在
            due2 = m2_by_invoice.loc[invoice, DUE_COL]
            collection = due1 - due2 * rate1
        else:  # ① 消失
            collection = due1 * rate1
        # ①② Refined Base Equiv = Base Equiv（来源行 M1）
        rows.append(_make_row(r, collection, wm, m1_name, r[BASE_EQUIV_COL]))

    # ③：仅在 M2 出现
    for _, r in m2.iterrows():
        if r[INVOICE_COL] in m1_invoices:
            continue
        collection = (r[PAID_COL] + r[PARTIAL_COL]) * r[RATE_COL]
        # ③ Refined Base Equiv = (Commission AR + Commission VAT) × Currency Rate
        refined = (r[AR_COL] + r[VAT_COL]) * r[RATE_COL]
        rows.append(_make_row(r, collection, wm, m2_name, refined))

    return rows


def age_debt_rows(
    age_debt: pd.DataFrame, m2: pd.DataFrame, year: int, month: int, wm: date
) -> list[dict[str, object]]:
    """④：Age_Debt 中 Invoice Date 属于 M2 当月、且 Invoice No 不在 M2 的发票（Report=ADR）。"""
    m2_invoices = set(m2[INVOICE_COL])
    mask = (
        (age_debt[INVOICE_DATE_COL].dt.year == year)
        & (age_debt[INVOICE_DATE_COL].dt.month == month)
    )
    rows: list[dict[str, object]] = []
    for _, r in age_debt[mask].iterrows():
        if r[INVOICE_COL] in m2_invoices:
            continue
        collection = (r[AR_COL] + r[VAT_COL]) * r[RATE_COL]
        # ④ Refined Base Equiv = Base Equiv（来源行 Age_Debt）
        rows.append(_make_row(r, collection, wm, ADR_REPORT, r[BASE_EQUIV_COL]))
    return rows


def add_split_columns(report: pd.DataFrame) -> pd.DataFrame:
    """存 Excel 前追加 4 列：按 Refined Base Equiv / Collection 的正负拆分。"""
    refined = report[REFINED_BASE_EQUIV_COL]
    collection = report[COLLECTION_COL]
    report[AR_OUT_COL] = refined.where(refined > 0, 0)
    report[AP_OUT_COL] = (-refined).where(refined < 0, 0)
    report[RECEIVED_COL] = collection.where(collection > 0, 0)
    report[PAID_AMOUNT_COL] = (-collection).where(collection < 0, 0)
    return report


def discover_month_files(input_dir: Path) -> list[tuple[int, int, Path]]:
    """扫描目录下所有 YYYYMM.xlsx，按 (年, 月) 排序。"""
    files: list[tuple[int, int, Path]] = []
    for p in input_dir.glob("*.xlsx"):
        name = p.stem
        if len(name) == 6 and name.isdigit() and 1 <= int(name[4:]) <= 12:
            files.append((int(name[:4]), int(name[4:]), p))
    files.sort(key=lambda t: (t[0], t[1]))
    return files


def build_report(input_dir: Path, age_debt_path: Path | None) -> pd.DataFrame:
    """对所有相邻月份对做比对并整合 Age_Debt，纵向堆叠。"""
    files = discover_month_files(input_dir)
    if len(files) < 2:
        raise SystemExit(f"目录 {input_dir} 下至少需要 2 个 YYYYMM.xlsx 文件，仅发现 {len(files)} 个。")

    age_debt = read_age_debt(age_debt_path) if age_debt_path else None

    all_rows: list[dict[str, object]] = []
    base_columns: list[str] | None = None
    for (_, _, p1), (y2, m2_month, p2) in zip(files, files[1:]):
        m1 = read_month(p1)
        m2 = read_month(p2)
        if base_columns is None:
            base_columns = m1.columns.tolist()
        wm = date(y2, m2_month, 1)
        all_rows.extend(compare_pair(m1, m2, wm, p1.stem, p2.stem))
        if age_debt is not None:
            all_rows.extend(age_debt_rows(age_debt, m2, y2, m2_month, wm))

    assert base_columns is not None
    row_columns = [
        *base_columns, COLLECTION_COL, WORKING_MONTH_COL, INCOME_TYPE_COL,
        REPORT_COL, REFINED_BASE_EQUIV_COL,
    ]
    report = pd.DataFrame(all_rows, columns=row_columns)
    return add_split_columns(report)


def main() -> None:
    parser = argparse.ArgumentParser(description="相邻月份月结报告比对（v2）")
    parser.add_argument(
        "-i", "--input-dir", type=Path, default=Path("."),
        help="存放 YYYYMM.xlsx 文件的目录（默认当前目录）",
    )
    parser.add_argument(
        "-o", "--output", type=Path, default=Path("collection_report.xlsx"),
        help="输出文件路径（默认 collection_report.xlsx）",
    )
    parser.add_argument(
        "--age-debt", type=Path, default=None,
        help=f"Age_Debt_Report 文件路径（默认在输入目录下找 {AGE_DEBT_FILENAME}）",
    )
    args = parser.parse_args()

    age_debt_path = args.age_debt or (args.input_dir / AGE_DEBT_FILENAME)
    if not age_debt_path.exists():
        print(f"⚠️  未找到 {age_debt_path}，跳过 Age_Debt 比对（④）。")
        age_debt_path = None

    report = build_report(args.input_dir, age_debt_path)
    report.to_excel(args.output, index=False)
    print(f"已生成 {args.output}，共 {len(report)} 行。")


if __name__ == "__main__":
    main()
