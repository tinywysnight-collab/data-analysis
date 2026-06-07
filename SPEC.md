# 月结报告相邻月份比对 — 需求规格（SPEC）

> 本文件是需求的唯一来源（single source of truth）。后续需求变更请直接更新本文件，
> 并同步调整 [monthly_compare.py](monthly_compare.py) 与 [selftest.py](selftest.py)。

- 版本：v4
- 最后更新：2026-06-06

---

## 1. 背景

财务每月产出一份月结报告。需要按相邻两个月做比对，计算每张发票本月实际的
回款 / 变动金额（Collection），标记工作月份（Working Month）与收入类型（Income Type），
并整合 Age_Debt_Report 中“当月即收款、财务报告里从未出现”的发票。

## 2. 输入

### 2.1 月度月结报告

| 项 | 说明 |
|---|---|
| 文件格式 | Excel（`.xlsx`） |
| 文件命名 | `YYYYMM.xlsx`（如 `202501.xlsx` … `202512.xlsx`、`202601.xlsx` …） |
| 文件范围 | **自动扫描**输入目录下所有 `YYYYMM.xlsx`，按时间排序，逐对相邻比对（含跨年） |
| 列结构 | 每个文件有数十列；**先裁剪、再去重**（见 2.3） |
| 原文件 | **只读不改** |

### 2.2 关键列（按列名引用，须落在保留的列里）

| 列名 | 含义 |
|---|---|
| `Invoice No` | 发票号（比对匹配键） |
| `Commission AR` | 应收佣金 |
| `Commission VAT` | 佣金税额（用于 Age_Debt 计算） |
| `Paid` | 已付 |
| `Partial Paid` | 部分支付 |
| `Currency` | 币种 |
| `Currency Rate` | 汇率 |
| `Total Due` | 应付/未结余额 |
| `Risk` | 风险标记（决定 Income Type） |
| `Invoice Date` | 开票日期（Age_Debt 用于按月筛选） |
| `Base Equiv` | 本位币等值（决定 Refined Base Equiv） |
| `Cover` | 与 `Ver` 合并为 `Cover Ver` |
| `Ver` | 版本号（源文件可能为 float） |

### 2.3 月度文件预处理（按顺序）

1. **去无用列**：保留第 **1–20、35、38** 列（按位置，1-indexed，共 22 列），保持原列顺序。
   上述关键列均落在保留的列里。
2. **去重**：按 `Invoice No` 分组，`Commission AR / Paid / Partial Paid` **求和**；
   其余列取**第一条**记录的值（假设同一发票其余列本来相同）；列顺序不变。
3. **合并 Cover Ver**：新增 `Cover Ver = Cover + "/" + Ver3`，其中 `Ver3` 为 `Ver`
   取整后补足 3 位（如 `5.0→005`、`12.0→012`；空值→`000`）。**保留**原 `Cover`/`Ver`，
   `Cover Ver` 插在 `Ver` 列之后。`Age_Debt`（④）做同样处理。

### 2.4 Age_Debt_Report

| 项 | 说明 |
|---|---|
| 文件名 | `Age_Debt_Report.xlsx`（默认在输入目录下查找；可用 `--age-debt` 指定） |
| 表头 | 前 3 行跳过，**第 4 行为列名行** |
| 列 | 列名与“保留后的列”一致 |
| `Invoice No` | 唯一（无需去重） |
| 缺失时 | 找不到文件则跳过 ④，仅输出 ①②③ |

## 3. 比对逻辑

对每一对相邻月份 `M1`（前月）、`M2`（后月）做比对。`X(M1)` / `X(M2)` 表示取自对应文件的行。
`Working Month` 统一为 **M2 年月的 1 号**（写成真正的日期值）。

| 情况 | 条件 | `Collection` 计算 | `Risk` 取自 | `Report` |
|---|---|---|---|---|
| ① 发票消失 | Invoice No 在 M1 有、M2 **无** | `Total Due(M1) × Currency Rate(M1)` | M1 | M1 文件名 |
| ② 发票仍在 | Invoice No 在 M1、M2 **都有** | `Total Due(M1) − Total Due(M2) × Currency Rate(M1)`（**先乘后减**） | M1 | M1 文件名 |
| ③ 发票新增 | Invoice No 在 M2 有、M1 **无** | `(Paid(M2) + Partial Paid(M2)) × Currency Rate(M2)` | M2 | M2 文件名 |
| ④ Age_Debt | `Invoice Date` 属于 **M2 当月** 且 Invoice No **不在 M2 文件** | `(Commission AR + Commission VAT) × Currency Rate` | Age_Debt | `ADR` |

`Report` 列标识每行来源：①② 取 M1 文件名、③ 取 M2 文件名、④ 为 `ADR`。文件名**不带扩展名**（如 `202501`）。

- ④ 仅依据 M2（每个月作为 M2 恰好出现一次）。**首月缺口不补**：列表中第一个文件永远只当 M1、
  不会当 M2，因此 Age_Debt 中开票日期属于“首月”的发票不会被处理。
- ④ 只看“不在 M2”，不额外排查 M1。

### 3.1 Income Type（新列，对每一行按其来源行的 `Risk`）

`Risk` 去首尾空格后**不区分大小写**：

- 前 3 字母 = `Fee` **或** 前 9 字母 = `Multiline` → `Income Type = Fee/NPT Fee`
- 否则 → `Income Type = Commission`

### 3.2 Refined Base Equiv（新列）

| 情况 | Refined Base Equiv |
|---|---|
| ①②④ | `Base Equiv`（来源行的值） |
| ③ | `(Commission AR + Commission VAT) × Currency Rate` |

### 3.3 存 Excel 前追加 4 列（按正负拆分）

| 列 | 规则 |
|---|---|
| `AR` | `Refined Base Equiv > 0` → 其值，否则 `0` |
| `AP` | `Refined Base Equiv < 0` → `-Refined Base Equiv`，否则 `0` |
| `Received` | `Collection > 0` → 其值，否则 `0` |
| `Paid Amount` | `Collection < 0` → `-Collection`，否则 `0` |

> 命名为 `Paid Amount` 而非 `Paid`，以免与关键列 `Paid` 混淆。值正好为 0 时对应两列均为 0。

## 4. 输出

| 项 | 说明 |
|---|---|
| 文件 | 单个 `.xlsx` |
| Sheet | **1 个 sheet** |
| 内容 | 所有相邻对的 ①②③ + ④ 结果**纵向堆叠**（靠 `Working Month` 区分来自哪对月份） |
| 每行列 | 保留后的原始列(含 `Cover`/`Ver`/`Cover Ver`) … + `Collection` + `Working Month` + `Income Type` + `Report` + `Refined Base Equiv` + `AR` + `AP` + `Received` + `Paid Amount` |
| 默认输出名 | `collection_report.xlsx` |

## 5. 运行方式

```bash
# 月度文件 + Age_Debt_Report.xlsx 都在当前目录：
uv run python monthly_compare.py

# 指定输入目录 / 输出文件 / Age_Debt 路径：
uv run python monthly_compare.py -i <输入目录> -o <输出.xlsx> --age-debt <路径>
```

新增月份（如未来的 2026 各月）直接把 `YYYYMM.xlsx` 放进输入目录即可，无需改代码。

## 6. 当前假设与已知边界（待确认/可扩展）

1. **数值列为纯数字**——`Total Due / Paid / Partial Paid / Currency Rate / Commission AR /
   Commission VAT` 不含货币符号、千分位逗号等文本；若含则需加清洗步骤。
2. **去重其余列同值**——同一发票的非求和列（`Total Due / Currency Rate / Risk` 等）取第一条；
   若同一发票多行这些值不一致，结果取第一条。
3. **关键列均落在保留的 22 列内**，且各文件列位置一致。
4. **Working Month / Invoice Date 为日期型**；如需文本格式（`2025/02/01`）需调整输出。
5. **Age_Debt 的列名与保留后的列一致**（含 `Invoice Date`、`Commission VAT`）。

## 7. 变更记录

| 版本 | 日期 | 变更 |
|---|---|---|
| v1 | 2026-06-06 | 初版：①②③ 比对规则、Working Month=M2 月 1 号、单文件单 sheet 堆叠输出 |
| v2 | 2026-06-06 | 列裁剪(1–20,35,38)+按 Invoice No 去重求和；自动发现 YYYYMM 文件(含跨年)；新增 Income Type(按 Risk)；新增 ④ Age_Debt 整合 |
| v3 | 2026-06-06 | 新增 Report 列：①②=M1 文件名、③=M2 文件名、④=ADR（文件名不带扩展名） |
| v4 | 2026-06-06 | 新增 Cover Ver 合并(Ver 补3位)；新增 Refined Base Equiv(③为公式、其余取 Base Equiv)；存盘前追加 AR/AP/Received/Paid Amount 正负拆分列 |
