---
lab:
  title: Charts and data visualization - Sales Performance Visualization Report
  module: Module 9 - Charts and Data Visualization
  description: In this lab, you build summary tables from the sales data and create column, bar, line, pie, doughnut, area, scatter and combo charts and sparklines, then learn which chart to choose for each question.
  duration: 90 minutes
  level: 200
  islab: true
  primarytopics:
    - Excel
---

# Sales Performance Visualization Report

**Lab Number:** 09
**Lab Title:** Create a Sales Performance Visualization Report
**Module:** Module 9 – Charts and Data Visualization
**Difficulty Level:** Intermediate
**Estimated Time:** 90 minutes
**Prerequisites:** Labs 03 and 04 (`SUMIFS`, tables).
**Excel version note:** Core chart types and sparklines work in Excel 2010+. The **Funnel, Map and Waterfall** chart types need Excel 2016+ (some only in 365). This lab uses only classic chart types.

## Problem Statement

Management cannot see patterns in tables of numbers. They need charts that show which regions, products and months drive sales, and which chart suits each question.

## Business Scenario

The Sales Director wants a one-page report with: sales by region, top products, monthly trend, category share, region-and-month patterns and profit margin by region.

## Learning Objectives

- Create column, bar, line, pie, doughnut, area, scatter and combo charts.
- Format titles, axes, legends and data labels.
- Add sparklines.
- Choose the right chart for the question.

## Dataset Download

| File | Records | Link |
|---|---|---|
| sales_transactions.csv | 1,200 | [Download](../Allfiles/Datasets/sales_transactions.csv) |

## Dataset Column Description

| Column | Description |
|---|---|
| OrderID, OrderDate | Order number and date (2024) |
| Region, Salesperson | 4 regions, 8 salespeople |
| Category, Product | 3 categories, 12 products |
| Quantity, UnitPrice, UnitCost, DiscountPct | Units, selling price, cost and discount % |

Calculated in Step 1: **NetSales**, **Profit**, **Month** (details in Lab 03).

## Step 1 – Import the dataset

1. Open **sales_transactions.csv**, save as **Sales-Charts.xlsx**, convert to table **tblSales**, and rename the sheet **Data**.
1. Add the columns from Lab 03: **NetSales** `=[@Quantity]*[@UnitPrice]*(1-[@DiscountPct]/100)`, **Profit** `=[@NetSales]-[@Quantity]*[@UnitCost]` and **Month** `=MONTH([@OrderDate])`.

## Step 2 – Inspect the data

1. Confirm 1,200 rows and that the **NetSales** total is **6,50,96,249.50**.
1. Add a sheet **Summary** for the tables used by charts. Charts should read from summary tables, not from raw data.

## Step 3 – Apply the Excel operation

Build these tables on **Summary** (use `SUMIFS`; fill each formula down):

**Table 1 – Region (A1:D5)**

| Region | Net Sales | Profit | Margin % |
|---|---|---|---|
| North | `=SUMIFS(tblSales[NetSales],tblSales[Region],A2)` | `=SUMIFS(tblSales[Profit],tblSales[Region],A2)` | `=C2/B2` |
| South, East, West | same | same | same |

**Table 2 – Product (F1:G13):** list the 12 products and `=SUMIFS(tblSales[NetSales],tblSales[Product],F2)`. Sort largest to smallest.

**Table 3 – Month (I1:J13):** months 1–12 and `=SUMIFS(tblSales[NetSales],tblSales[Month],I2)`.

**Table 4 – Category (L1:M4):** Electronics, Furniture, Office Supplies and their net sales.

**Table 5 – Region × Month (A20:M24):** regions in A21:A24, months 1–12 in B20:M20, formula `=SUMIFS(tblSales[NetSales],tblSales[Region],$A21,tblSales[Month],B$20)` (note the **mixed references**).

## Step 4 – Apply formulas

Check values before charting:

| Region | Net Sales | Profit | Margin |
|---|---|---|---|
| North | 1,46,06,054.50 | 29,37,504.50 | 20.1% |
| South | 1,92,86,136.00 | 39,02,756.00 | 20.2% |
| East | 1,40,88,106.50 | 28,33,206.50 | 20.1% |
| West | 1,71,15,952.50 | 34,06,322.50 | 19.9% |

## Step 5 – Analyse the results

Create each chart below. For each one: select the data, **Insert > Charts**, set a title, remove clutter and add data labels where useful.

| # | Chart | Data | Insert command | Use it when |
|---|---|---|---|---|
| 1 | **Clustered column** | Region and Net Sales | Insert > Column > Clustered Column | Comparing a few categories |
| 2 | **Bar** | Top 6 products | Insert > Bar > Clustered Bar | Long category names or rankings; sort largest at top (Format Axis > Categories in reverse order) |
| 3 | **Line with markers** | Month and Net Sales | Insert > Line > Line with Markers | Trends over time |
| 4 | **Pie** | Category and Net Sales | Insert > Pie | Share of a total, 2–5 slices only |
| 5 | **Doughnut** | Region share | Insert > Pie > Doughnut | Same as pie, with a KPI in the centre |
| 6 | **Area** | Month and Net Sales | Insert > Area | Volume over time, emphasising magnitude |
| 7 | **Scatter (XY)** | Quantity (X) vs NetSales (Y) from `Data` for 200 rows | Insert > Scatter | Relationship between two numeric variables |
| 8 | **Combo** | Region: Net Sales (columns) and Margin % (line, secondary axis) | Insert > Combo > Clustered Column - Line on Secondary Axis | Two measures with different scales |

Formatting checklist for every chart:

1. **Chart Title:** meaningful, e.g. *South leads regional net sales*.
1. **Axis titles** and number format: right-click axis > **Format Axis > Number > Display Units: Lakhs** (or Millions) to avoid long numbers.
1. **Legend:** delete if there is one series; place at the bottom if needed.
1. **Data Labels:** add only to the points that matter.
1. **Gridlines:** make light grey or remove.
1. **Colour:** use one colour for single series; highlight the key bar in a different colour.

![Chart formatting checklist applied](Linked_image_Files/lab09_step5_formatted-column.png)

### Sparklines

1. Click cell **N21** next to the Region × Month table.
1. **Insert > Sparklines > Line**, Data Range **B21:M21**, Location **N21**, **OK**.
1. Fill down to **N24**. Turn on **High Point** and **Low Point** in **Sparkline Design**.

    > *Sparklines are tiny charts inside a cell: perfect for showing a trend beside each row of a table.*

## Step 6 – Create visualization

Arrange the charts on a new sheet named **Report** in a 2 × 3 grid:

1. Copy each chart (**Ctrl + C**) and paste on **Report**.
1. Resize with **Alt + drag** to snap to cell borders.
1. Align using **Chart Format > Align**.
1. Give each chart the same font and colour palette (**Page Layout > Colors**).

    ![Sales performance report layout](Linked_image_Files/lab09_step6_report-layout.png)

### Choosing the right chart

| Question | Best chart |
|---|---|
| Which region sold the most? | Column |
| Which products rank highest? | Bar (sorted) |
| How did sales change through the year? | Line |
| What share does each category hold? | Pie / Doughnut (few categories) |
| Is higher quantity linked with higher sales? | Scatter |
| How do revenue and margin compare together? | Combo |
| Trend per region inside a table | Sparklines |

> ***Avoid*:** 3D charts, pies with many slices, truncated axes that exaggerate differences, and more than 6 colours.

## Expected Output

| Chart | What you should see |
|---|---|
| Region column | South tallest (₹1.93 crore), East shortest (₹1.41 crore) |
| Product bar | **Laptop** first (₹2,89,63,000), then **Smartphone** (₹1,18,15,100), **Standing Desk** (₹84,78,000) |
| Monthly line | Low in **Jan (₹41.6 lakh)**, peak in **Dec (₹73.3 lakh)** |
| Category pie | **Electronics 72.5%**, Furniture 24.7%, Office Supplies 2.7% |
| Combo | Margins all close to 20% (19.9% to 20.2%) |
| Sparklines | Four mini lines with high and low points marked |

## Screenshot of final result

![Final sales visualization report](Linked_image_Files/lab09_final_report.png)

## Business Insights

- Electronics generate 72.5% of net sales, so the business depends heavily on laptops and smartphones.
- Regions have similar margins (about 20%); the difference between them is **volume**, not profitability.
- December is the seasonal peak; January and April are weak months for campaigns.

## Practice Questions

1. Why should you avoid a pie chart with 10 categories?
1. When do you use a combo chart?
1. What does `$A21` versus `B$20` mean in the Region × Month formula?
1. How do you change numbers on an axis to lakhs?
1. What is a sparkline?

<details>
<summary>Answer key</summary>

1. Slices become tiny and hard to compare; use a sorted bar chart.
1. When two measures have different units or scales (e.g. amount and percentage).
1. `$A21` locks the column (region list) and `B$20` locks the row (month headers).
1. Format Axis > Number > Display Units, or a custom format.
1. A tiny chart inside a cell showing a trend.

</details>

## Lab Assignment

Add a **Salesperson ranking bar chart** (8 salespeople) and a **profit margin by category column chart**. Write a 5-line summary for the Sales Director with one recommendation based on the charts. Submit the workbook.
