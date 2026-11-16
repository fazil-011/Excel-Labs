---
lab:
  title: Basic formulas and functions - Monthly Sales Summary
  module: Module 3 - Basic Formulas and Functions
  description: In this lab, you use SUM, AVERAGE, COUNT, COUNTA, MAX, MIN, ROUND, SUBTOTAL and AutoSum on a 1,200-row sales dataset to build a monthly sales summary report.
  duration: 60 minutes
  level: 100
  islab: true
  primarytopics:
    - Excel
---

# Monthly Sales Summary Report

**Lab Number:** 03
**Lab Title:** Create a Monthly Sales Summary Report
**Module:** Module 3 – Basic Formulas and Functions
**Difficulty Level:** Beginner
**Estimated Time:** 60 minutes
**Prerequisites:** Labs 01 and 02 (references, tables).

## Problem Statement

The sales team has 1,200 order lines for 2024. Management wants total sales, profit, order counts and a month-by-month summary, but only raw transactions exist.

## Business Scenario

You are a junior analyst. The Sales Director asks: *What were our net sales, cost and profit for 2024? What is the average and the largest order? How did each month perform?*

## Learning Objectives

- Write formulas with arithmetic operators and correct operator precedence.
- Use `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MAX`, `MIN`, `ROUND`, `SUBTOTAL` and AutoSum.
- Use structured table references.
- Explain why `SUBTOTAL` is better than `SUM` for filtered data.

## Dataset Download

| File | Records | Link |
|---|---|---|
| sales_transactions.csv | 1,200 | [Download](../Allfiles/Datasets/sales_transactions.csv) |

## Dataset Column Description

| Column | Type | Description |
|---|---|---|
| OrderID | Text | ORD-10001 to ORD-11200 |
| OrderDate | Date | 1 Jan 2024 to 31 Dec 2024 |
| Region | Text | North, South, East, West |
| Salesperson | Text | 8 salespeople (2 per region) |
| Category | Text | Electronics, Furniture, Office Supplies |
| Product | Text | 12 products |
| Quantity | Whole number | Units sold (1–10) |
| UnitPrice | Number | Price per unit in INR |
| UnitCost | Number | Cost per unit in INR |
| DiscountPct | Number | Discount percentage (0, 5, 10, 15) |

## Step 1 – Import the dataset

1. Open **sales_transactions.csv** in Excel and save as **Sales-Summary.xlsx**.
1. Click a data cell, press **Ctrl + T**, confirm headers, and name the table **tblSales**.
1. Rename the sheet **Data**. Make sure **OrderDate** shows as a date (**Ctrl + 1 > Date**).

## Step 2 – Inspect the data

1. Select **Quantity** (G2:G1201). The status bar should show **Count: 1200**.
1. Check there are no blanks: select **A1:J1201**, press **F5 > Special > Blanks**. You should get *"No cells were found"*.
1. Scroll to the last row to confirm **ORD-11200** is the last order.

## Step 3 – Apply the Excel operation

1. Add four new columns to the table by typing a header in **K1**, **L1**, **M1**, **N1**:

    | Header | Meaning |
    |---|---|
    | GrossSales | Quantity × UnitPrice |
    | DiscountAmount | GrossSales × DiscountPct / 100 |
    | NetSales | GrossSales − DiscountAmount |
    | TotalCost | Quantity × UnitCost |

1. Add **O1: Profit** (NetSales − TotalCost) and **P1: Month** (month number of OrderDate).
1. Press **Ctrl + S**.

## Step 4 – Apply formulas

Type each formula in the first data row; the table fills the rest automatically.

| Column | Formula |
|---|---|
| GrossSales (K2) | `=[@Quantity]*[@UnitPrice]` |
| DiscountAmount (L2) | `=[@GrossSales]*[@DiscountPct]/100` |
| NetSales (M2) | `=[@GrossSales]-[@DiscountAmount]` |
| TotalCost (N2) | `=[@Quantity]*[@UnitCost]` |
| Profit (O2) | `=[@NetSales]-[@TotalCost]` |
| Month (P2) | `=MONTH([@OrderDate])` |

> ***Why*:** Excel follows **PEMDAS** order: brackets, exponents, multiplication/division, addition/subtraction. `=A1+B1*C1` multiplies first. Use brackets to control order.

Check row 2 (ORD-10001): Quantity 1 × ₹7,500 = 7,500; 5% discount = 375; NetSales = **7,125**; cost = 5,200; Profit = **1,925**.

### Summary sheet

1. Add a new sheet named **Summary**. Type the labels in **A1:A10** (in the order below, so Total Net Sales is in B4 and Total Profit in B6) and the formulas in column B:

    | Label | Formula | Why this function |
    |---|---|---|
    | Total Orders | `=COUNTA(tblSales[OrderID])` | Counts non-empty text cells |
    | Orders with a Quantity | `=COUNT(tblSales[Quantity])` | Counts numeric cells only |
    | Total Units Sold | `=SUM(tblSales[Quantity])` | Adds values |
    | Total Net Sales | `=SUM(tblSales[NetSales])` | Adds values |
    | Total Cost | `=SUM(tblSales[TotalCost])` | Adds values |
    | Total Profit | `=SUM(tblSales[Profit])` | Adds values |
    | Average Order Value | `=ROUND(AVERAGE(tblSales[NetSales]),2)` | Mean; ROUND to 2 decimals |
    | Largest Order | `=MAX(tblSales[NetSales])` | Highest value |
    | Smallest Order | `=MIN(tblSales[NetSales])` | Lowest value |
    | Profit Margin % | `=B6/B4` (Profit ÷ Net Sales) | Ratio of two totals |

1. Format Net Sales, Cost, Profit, Average, Largest and Smallest as **₹ Currency**; format the margin as **0.0%**.
1. **AutoSum:** click an empty cell below the Quantity column on **Data** (outside the table) and select **Home > AutoSum (Σ)**. Excel inserts `SUM` and guesses the range. Always check the highlighted range before pressing **Enter**.

## Step 5 – Analyse the results

### Monthly summary using SUBTOTAL

1. On **Data**, below the table (for example **M1204**) type:

    ```excel
    =SUBTOTAL(109,tblSales[NetSales])
    ```

    > *`109` means SUM that ignores rows hidden by a filter. `9` also sums but includes manually hidden rows.*

1. Click the filter arrow on **Month** and tick only **1** (January). The SUBTOTAL cell shows the January net sales. Compare with `SUM`: it still shows the full total.
1. On the **Summary** sheet, build a table in **D1:F13** with headers **Month**, **Net Sales**, **Profit**. Filter month by month and type each value (or copy > Paste Values) into the table.

    > ***Shortcut*:** After Lab 04 you can replace this manual process with `=SUMIFS(tblSales[NetSales],tblSales[Month],D2)`.

1. Add a **Total** row with `=SUM(E2:E13)` and confirm it equals Total Net Sales.

## Step 6 – Create visualization

1. Select **D1:E13** and choose **Insert > Charts > Line with Markers**.
1. Title the chart **Monthly Net Sales 2024**, and add axis title **₹**.

    ![Monthly net sales line chart](Linked_image_Files/lab03_step6_monthly-line.png)

## Expected Output

| Metric | Expected value |
|---|---|
| Total Orders | 1,200 |
| Total Units Sold | 6,589 |
| Total Net Sales | ₹6,50,96,249.50 |
| Total Profit | ₹1,30,79,789.50 |
| Profit Margin | 20.1% |
| Average Order Value | ₹54,246.87 |
| Largest Order | ₹5,50,000 |
| Smallest Order | ₹153 |

Monthly net sales (rounded to the nearest rupee):

| Month | Net Sales | Month | Net Sales |
|---|---|---|---|
| Jan | 41,63,798 | Jul | 62,81,609 |
| Feb | 62,28,702 | Aug | 51,98,936 |
| Mar | 48,74,738 | Sep | 54,48,317 |
| Apr | 43,91,168 | Oct | 58,25,824 |
| May | 54,83,831 | Nov | 51,29,176 |
| Jun | 47,41,541 | Dec | 73,28,612 |

## Screenshot of final result

![Final summary sheet and chart](Linked_image_Files/lab03_final_summary.png)

## Business Insights

- December is the peak month (₹73.3 lakh) and January the weakest (₹41.6 lakh): plan stock and staffing for Q4.
- The overall margin is about 20%. Check which products drag it down in Labs 10 and 13.
- Average order value is ₹54,247, but the smallest order is ₹153: orders vary hugely, so the average alone is misleading.

## Practice Questions

1. What is the difference between `COUNT` and `COUNTA`?
1. Why use `SUBTOTAL(109, ...)` instead of `SUM` on a filtered table?
1. What does `=ROUND(1234.5678,-2)` return?
1. Write a formula for the average quantity per order.
1. What is the result of `=2+3*4` and `=(2+3)*4`?

<details>
<summary>Answer key</summary>

1. COUNT counts numbers; COUNTA counts non-empty cells of any type.
1. SUBTOTAL ignores rows hidden by filters; SUM includes them.
1. 1200.
1. `=AVERAGE(tblSales[Quantity])` (5.49).
1. 14 and 20.

</details>

## Lab Assignment

Create a one-page **Regional Summary** with total net sales for each region using only filters and `SUBTOTAL(109, ...)`. Record the four results, then check them against Lab 04 using `SUMIFS`. Submit the workbook with a 3-line note about which method was faster and why.
