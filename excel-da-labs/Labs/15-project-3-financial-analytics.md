---
lab:
  title: Project 3 - Financial Analytics
  module: Module 13 - Real-World Projects
  description: In this capstone project, you build a monthly profit and loss statement, compare budget with actual results, calculate financial KPIs and create a financial dashboard.
  duration: 150 minutes
  level: 300
  islab: true
  primarytopics:
    - Excel
---

# Project 3 – Financial Analytics

**Lab Number:** 15
**Lab Title:** Financial Analytics – P&L and Budget vs Actual
**Module:** Module 13 – Real-World Projects
**Difficulty Level:** Advanced
**Estimated Time:** 150 minutes
**Prerequisites:** Labs 04, 09, 10, 12.
**Excel version note:** The **Waterfall** chart (Insert > Waterfall) needs Excel 2016 or later (Windows/365). If unavailable, build a stacked-column waterfall manually.

## Problem Statement

Finance reports revenue and expenses in a long list of 96 monthly entries. The CFO cannot quickly see monthly profit, whether the company met its budget or which expense lines overran.

## Business Scenario

The CFO asks: *What was our revenue, cost and profit for 2024? Which months missed budget? Which lines are over or under budget? What are our margins?*

## Learning Objectives

- Structure a profit and loss (P&L) statement with `SUMIFS`.
- Calculate budget variance and variance %, and understand favourable vs unfavourable.
- Calculate gross margin and net margin.
- Build a financial KPI dashboard.

## Dataset Download

| Dataset | Link |
|---|---|
| **Practice dataset (provided, synthetic):** finance_monthly.csv, 96 rows | [Download](../Allfiles/Datasets/finance_monthly.csv) |
| **Real-world option:** Financial Sample (Kaggle) | https://www.kaggle.com/datasets/atharvaarya25/financials |

> *The Kaggle financial sample is a sales and profit dataset (by segment, country, product, discount band). It is not a P&L with budget columns, so use it for a **sales-and-profit analysis** after this project. It is not included in this package: download it yourself from Kaggle.*

## Dataset Column Description (Data dictionary)

| Column | Type | Description |
|---|---|---|
| Month | Date | First day of the month, Jan–Dec 2024 |
| Account | Text | Product Sales, Service Revenue, Cost of Goods Sold, Salaries, Rent, Marketing, Utilities, Other Expenses |
| AccountType | Text | Revenue, COGS, Opex |
| Budget | Number | Planned amount for the month (INR) |
| Actual | Number | Actual amount for the month (INR) |

## Required Excel functions

`SUMIFS`, `IF`, `IFERROR`, `TEXT`, `EOMONTH`, `ROUND`, PivotTables, Conditional Formatting (icon sets), Combo and Waterfall charts.

## Step 1 – Import the dataset

1. Open **finance_monthly.csv**, save as **Finance-Analytics.xlsx**, convert to table **tblFin**, rename the sheet **Data**.
1. Format **Month** as `mmm-yy` and Budget/Actual as ₹.

## Step 2 – Inspect the data

1. Count: 96 rows = 12 months × 8 accounts.
1. Pivot **AccountType × Month** to confirm each month has 2 revenue, 1 COGS and 5 Opex lines.
1. Check totals: Actual revenue **₹1,57,81,207**; Budget revenue **₹1,53,36,000**.

## Step 3 – Apply the Excel operation

Add helper columns to **tblFin**:

| Column | Formula |
|---|---|
| Variance | `=[@Actual]-[@Budget]` |
| Favourable | `=IF([@AccountType]="Revenue",IF([@Variance]>=0,"Yes","No"),IF([@Variance]<=0,"Yes","No"))` |
| MonthNo | `=MONTH([@Month])` |

> *For revenue, **higher than budget** is favourable. For costs, **lower than budget** is favourable. Many beginners get this wrong.*

## Step 4 – Apply formulas

Create a **P&L** sheet with months across (B1:M1 = 1 to 12, N1 = Total) and rows:

| Row | Formula (cell B for month 1, copy across) |
|---|---|
| Revenue | `=SUMIFS(tblFin[Actual],tblFin[AccountType],"Revenue",tblFin[MonthNo],B$1)` |
| COGS | `=SUMIFS(tblFin[Actual],tblFin[AccountType],"COGS",tblFin[MonthNo],B$1)` |
| **Gross Profit** | `=B2-B3` |
| Operating Expenses | `=SUMIFS(tblFin[Actual],tblFin[AccountType],"Opex",tblFin[MonthNo],B$1)` |
| **Net Profit** | `=B4-B5` |
| Gross Margin % | `=IFERROR(B4/B2,0)` |
| Net Margin % | `=IFERROR(B6/B2,0)` |
| Budget Net Profit | `=SUMIFS(tblFin[Budget],tblFin[AccountType],"Revenue",tblFin[MonthNo],B$1)-SUMIFS(tblFin[Budget],tblFin[AccountType],"COGS",tblFin[MonthNo],B$1)-SUMIFS(tblFin[Budget],tblFin[AccountType],"Opex",tblFin[MonthNo],B$1)` |
| Profit Variance | `=B6-B9` |

Column **N** = `=SUM(B2:M2)` for amounts and recalculated ratios for margins.

![P&L layout](Linked_image_Files/p3_step4_pl-layout.png)

## Step 5 – Analyse the results

Create a **Budget vs Actual** table by account (Pivot: Account in Rows; Budget, Actual, Variance in Values; add calculated field **Var%** = `Variance/Budget`).

| Item | Budget | Actual | Variance | Result |
|---|---|---|---|---|
| Revenue | 1,53,36,000 | 1,57,81,207 | +4,45,207 (+2.9%) | Favourable |
| COGS | 66,41,000 | 66,80,392 | +39,392 (+0.6%) | **Unfavourable** |
| Operating expenses | 57,63,000 | 56,67,290 | −95,710 (−1.7%) | Favourable |
| **Net profit** | 29,32,000 | **34,33,525** | **+5,01,525 (+17.1%)** | Favourable |

By account: Product Sales +3,92,294; Service Revenue +52,913; Salaries −1,02,845 (under budget); Other Expenses +13,020 (over); Marketing −6,489; Rent −45; Utilities +649.

Key ratios (full year, actual): **Gross profit ₹91,00,815 (57.7%)**, **Net margin 21.8%**.

Monthly net profit (actual): Jan 3,68,113; Feb 1,93,323; Mar 2,78,091; Apr 3,50,657; May 2,82,307; Jun 2,69,338; Jul 2,20,127; Aug 3,12,858; **Sep 4,00,093 (best)**; **Oct 1,54,883 (worst)**; Nov 3,22,822; Dec 2,80,913. Months below budget profit: **Feb, Jul, Oct**.

## Step 6 – Create visualization

1. **Combo chart:** Net Profit (column) and Net Margin % (line, secondary axis) by month.
1. **Clustered column:** Budget vs Actual net profit by month.
1. **Waterfall (365/2016+):** Revenue → COGS → Gross Profit → Opex → Net Profit.
1. **Conditional formatting:** variance column with green up arrow / red down arrow icon sets (**Home > Conditional Formatting > Icon Sets**). For costs, reverse the icon order.
1. KPI cards: Revenue, Gross Profit, Net Profit, Net Margin %, Budget Achievement % (`Actual NP / Budget NP`).

    ![Financial KPI dashboard](Linked_image_Files/p3_step6_dashboard.png)

## Expected Output

| KPI | Value |
|---|---|
| Revenue (actual) | ₹1,57,81,207 |
| Gross profit | ₹91,00,815 (57.7%) |
| Net profit | ₹34,33,525 (21.8%) |
| Budget net profit | ₹29,32,000 |
| Budget achievement | 117.1% |

## Screenshot of final result

![Final financial dashboard](Linked_image_Files/p3_final_dashboard.png)

## Business Insights

1. Net profit beat the budget by 17.1%, driven by revenue (+2.9%) and lower salary cost.
1. COGS overran budget while revenue rose, so gross margin per rupee is slightly below plan; review supplier prices.
1. October had the lowest profit (₹1.55 lakh) because COGS and opex peaked; investigate one-off costs.
1. Three months (Feb, Jul, Oct) missed budget: forecast cash for these months more conservatively.

## Interview questions based on this project

1. What is the difference between gross profit and net profit?
1. How do you decide whether a variance is favourable for revenue vs cost?
1. How would you build this P&L if new accounts are added next year?
1. What is the advantage of `SUMIFS` over manual linking?
1. How would you present the dashboard to a non-finance manager?

## Practice Questions

1. If Actual revenue is 110 and Budget is 100, what is variance % and is it favourable?
1. If Actual rent is 95 and Budget is 100, is the variance favourable?
1. What is the formula for Gross Margin %?
1. Why use `IFERROR` in margin formulas?
1. Which chart best shows how revenue turns into profit?

<details>
<summary>Answer key</summary>

1. +10%, favourable.
1. Yes, a cost lower than budget is favourable.
1. Gross Profit ÷ Revenue.
1. To avoid `#DIV/0!` when revenue is 0.
1. Waterfall chart.

</details>

## Lab Assignment

Add a **forecast**: use `FORECAST.LINEAR` (Excel 2016+) on monthly revenue to estimate January–March 2025 and add it to the chart. Write a 5-line commentary for the CFO. Optional: analyse the Kaggle Financial Sample by segment and country.
