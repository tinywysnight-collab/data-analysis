"""用合成数据自测 v2 各项逻辑：列裁剪、去重、①②③④、Income Type。"""

from datetime import date

import pandas as pd

from monthly_compare import (
    COLLECTION_COL,
    INCOME_TYPE_COL,
    REPORT_COL,
    WORKING_MONTH_COL,
    age_debt_rows,
    compare_pair,
    dedup_invoices,
    income_type,
    select_kept_columns,
)

KEPT = [
    "Invoice No", "Commission AR", "Commission VAT", "Paid", "Partial Paid",
    "Currency", "Currency Rate", "Total Due", "Risk", "Invoice Date",
]


def row(inv, *, ar=0, vat=0, paid=0, partial=0, rate=1.0, due=0,
        risk="Normal", inv_date=None, cur="USD"):
    return {
        "Invoice No": inv, "Commission AR": ar, "Commission VAT": vat,
        "Paid": paid, "Partial Paid": partial, "Currency": cur,
        "Currency Rate": rate, "Total Due": due, "Risk": risk,
        "Invoice Date": inv_date,
    }


# ---- Income Type ----
assert income_type("Fee xxx") == "Fee/NPT Fee"
assert income_type("  fEe") == "Fee/NPT Fee"           # 去空格 + 不区分大小写
assert income_type("Multiline abc") == "Fee/NPT Fee"
assert income_type("MULTILINE") == "Fee/NPT Fee"
assert income_type("Normal") == "Commission"
assert income_type(None) == "Commission"
print("Income Type ✅")

# ---- 去重：SUM_COLS 求和，其余取第一条，列顺序不变 ----
dup = pd.DataFrame([
    row("A", ar=10, paid=1, partial=2, due=100, rate=7.0, risk="Fee"),
    row("A", ar=5, paid=3, partial=4, due=999, rate=9.9, risk="zzz"),  # 其余列取第一条
    row("B", ar=8, paid=0, partial=0, due=50, rate=6.0),
])
deduped = dedup_invoices(dup)
assert list(deduped.columns) == list(dup.columns)        # 列顺序不变
a = deduped.set_index("Invoice No").loc["A"]
assert a["Commission AR"] == 15 and a["Paid"] == 4 and a["Partial Paid"] == 6  # 求和
assert a["Total Due"] == 100 and a["Currency Rate"] == 7.0 and a["Risk"] == "Fee"  # 第一条
assert len(deduped) == 2
print("Dedup ✅")

# ---- 列裁剪：保留位置 1-20,35,38 ----
wide = pd.DataFrame([[i for i in range(40)]], columns=[f"c{i}" for i in range(40)])
kept = select_kept_columns(wide)
assert list(kept.columns) == [f"c{i}" for i in range(20)] + ["c34", "c37"]
print("Select columns ✅")

# ---- ①②③ ----
m1 = pd.DataFrame([
    row("INV1", due=100, rate=7.0, risk="Normal"),         # 仍在
    row("INV2", due=50, rate=6.0, risk="Fee deal"),         # 消失
])
m2 = pd.DataFrame([
    row("INV1", due=40, rate=7.2),                          # 仍在
    row("INV3", paid=10, partial=5, rate=6.5, risk="Multiline x"),  # 新增
])
wm = date(2025, 2, 1)
rows = compare_pair(m1, m2, wm, "202501", "202502")
by = pd.DataFrame(rows).set_index("Invoice No")
assert by.loc["INV1", COLLECTION_COL] == 100 - 40 * 7.0           # ② 先乘后减 = -180
assert by.loc["INV2", COLLECTION_COL] == 50 * 6.0                 # ① = 300
assert by.loc["INV3", COLLECTION_COL] == (10 + 5) * 6.5           # ③ = 97.5
assert by.loc["INV2", INCOME_TYPE_COL] == "Fee/NPT Fee"          # Fee 前缀
assert by.loc["INV3", INCOME_TYPE_COL] == "Fee/NPT Fee"          # Multiline 前缀
assert by.loc["INV1", INCOME_TYPE_COL] == "Commission"
assert (pd.DataFrame(rows)[WORKING_MONTH_COL] == wm).all()
# Report：①② 取 M1 名，③ 取 M2 名
assert by.loc["INV1", REPORT_COL] == "202501"                    # ②
assert by.loc["INV2", REPORT_COL] == "202501"                    # ①
assert by.loc["INV3", REPORT_COL] == "202502"                    # ③
print("Compare ①②③ + Report ✅")

# ---- ④ Age_Debt ----
age = pd.DataFrame([
    row("AD1", ar=100, vat=20, rate=7.0, risk="Fee", inv_date=pd.Timestamp("2025-02-15")),  # 命中
    row("AD2", ar=50, vat=5, rate=6.0, inv_date=pd.Timestamp("2025-01-31")),  # 非当月，排除
    row("INV1", ar=1, vat=1, rate=1.0, inv_date=pd.Timestamp("2025-02-10")),  # 当月但在 M2，排除
])
ad_rows = age_debt_rows(age, m2, 2025, 2, wm)
assert len(ad_rows) == 1
ad = ad_rows[0]
assert ad["Invoice No"] == "AD1"
assert ad[COLLECTION_COL] == (100 + 20) * 7.0                     # = 840
assert ad[INCOME_TYPE_COL] == "Fee/NPT Fee"
assert ad[WORKING_MONTH_COL] == wm
assert ad[REPORT_COL] == "ADR"                                   # ④ Report
print("Age_Debt ④ ✅")

print("\n全部自测通过 ✅")
