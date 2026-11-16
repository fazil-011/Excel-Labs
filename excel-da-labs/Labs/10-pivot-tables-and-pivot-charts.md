---
lab:
  title: Pivot tables and pivot charts - Interactive Sales Analysis Report
  module: Module 10 - Pivot Tables and Pivot Charts
  description: In this lab, you build pivot tables, a calculated field, slicers, a timeline and pivot charts to answer five business questions from 1,200 sales transactions.
  duration: 90 minutes
  level: 200
  islab: true
  primarytopics:
    - Excel
---

# Interactive Sales Analysis Report

**Lab Number:** 10
**Lab Title:** Create an Interactive Sales Analysis Report
**Module:** Module 10 – Pivot Tables and Pivot Charts
**Difficulty Level:** Intermediate
**Estimated Time:** 90 minutes
**Prerequisites:** Labs 03, 04 and 09.
**Excel version note:** PivotTables, slicers and PivotCharts work in Excel 2013+. **Timelines** need Excel 2013+. Automatic date grouping into months/quarters/years is available in all versions; Excel 365 groups dates automatically when you drag them.

## Problem Statement

Managers keep asking new questions of the same data: by region, product, person or month. Writing a new formula for each question is slow. A PivotTable lets users summarise 1,200 rows in seconds.

## Business Scenario

Answer these five business questions:

1. What are the total sales by region?
1. Which product has the highest revenue?
1. Which salesperson performs best?
1. What are the monthly sales trends?
1. Which category generates the highest profit?

## Learning Objectives

- Create PivotTables and place fields in Rows, Columns, Values and Filters.
- Group dates, sort and filter pivots.
- Create a calculated field and change value field settings.
- Add slicers and a timeline and connect them to several pivots.
- Create PivotCharts and refresh pivots.

## Dataset Download

| File | Records | Link |
|---|---|---|
| sales_transactions.csv | 1,200 | [Download](../Allfiles/Datasets/sales_transactions.csv) |

## Dataset Column Description

See Lab 03. Add the calculated columns in Step 1.

## Step 1 – Import the dataset

1. Open **sales_transactions.csv**, save as **Sales-Pivot.xlsx** and convert to table **tblSales**.
1. Add columns **NetSales**, **Profit** and **Month** (see Lab 09 Step 1).
1. Confirm there are **no blank rows and no merged cells** and that each column has one header. Pivot tables fail on clean-looking but messy data.

## Step 2 – Inspect the data

1. Check that **OrderDate** holds real dates: select the column and the status bar should show **Count** and **Sum**; a Sum means numbers (dates).
1. Total **NetSales** should be **6,50,96,249.50** and **Profit 1,30,79,789.50**.

## Step 3 – Apply the Excel operation

### Create the first PivotTable

1. Click any cell in **tblSales**, select **Insert > PivotTable**.
1. Choose **New Worksheet** and select **OK**. Rename the sheet **PT_Region**.
1. In the **PivotTable Fields** pane:
    - Drag **Region** to **Rows**.
    - Drag **NetSales** to **Values**.

    ![PivotTable field list with Region in Rows and NetSales in Values](Linked_image_Files/lab10_step3_field-list.png)

1. Select any value cell, then **PivotTable Analyze > Field Settings > Number Format > Currency** or a **Number** format with 0 decimals.

> ***Understand the four areas:*** **Rows** = categories down the side. **Columns** = categories across the top. **Values** = numbers to calculate. **Filters** = a drop-down to limit the whole report.

## Step 4 – Apply formulas

### Part A – Value field settings

1. Right-click a value > **Value Field Settings**:
    - **Summarize Values By:** Sum, Count, Average, Max, Min.
    - **Show Values As:** **% of Grand Total** to show share.
1. Add **NetSales** again to Values, set **Show Values As = % of Grand Total** and rename it **% Share**.

### Part B – Calculated field

1. Click inside the pivot and select **PivotTable Analyze > Fields, Items & Sets > Calculated Field**.
1. Name: **Margin%**. Formula: `=Profit/NetSales`. Click **Add > OK**.
1. Format the new column as **0.0%**.

    > *A calculated field works on **sums** of the underlying fields. `Profit/NetSales` therefore gives the correct ratio of totals, not an average of row percentages.*

### Part C – Group dates

1. Create a second pivot on a new sheet **PT_Month**.
1. Drag **OrderDate** to **Rows**. If Excel does not group, right-click a date > **Group > Months and Quarters** (select both) > **OK**.
1. Drag **NetSales** to **Values**.

### Part D – Sorting and filtering

1. On **PT_Region**, click the **Row Labels** arrow > **More Sort Options > Descending by NetSales**.
1. On a **Product** pivot (new sheet **PT_Product**), drag **Product** to **Rows** and **NetSales** to **Values**. Click the arrow > **Value Filters > Top 10** > show **Top 5** items by **NetSales**.

## Step 5 – Analyse the results

Build one pivot per question (new sheet each, or all on a **Pivots** sheet):

| Question | Rows | Values | Sort / filter |
|---|---|---|---|
| 1. Sales by region | Region | NetSales | Descending |
| 2. Highest-revenue product | Product | NetSales | Descending |
| 3. Best salesperson | Salesperson | NetSales | Descending |
| 4. Monthly trend | OrderDate (Months) | NetSales | Chronological |
| 5. Most profitable category | Category | Profit | Descending |

### Add slicers and a timeline

1. Click inside **PT_Region**, **PivotTable Analyze > Insert Slicer**, tick **Category** and **Region**, **OK**.
1. **PivotTable Analyze > Insert Timeline**, tick **OrderDate**, **OK**. Use the timeline to select **Jul – Dec**.
1. To drive **all pivots** from one slicer: right-click the slicer > **Report Connections** > tick every pivot (PT_Region, PT_Product, PT_Salesperson, PT_Month, PT_Category) > **OK**.

    ![Slicer connected to multiple pivots](Linked_image_Files/lab10_step5_report-connections.png)

### Refreshing

1. Change one **Quantity** in the **tblSales** table. The pivots do not update automatically.
1. Select **PivotTable Analyze > Refresh** (or **Alt + F5**; **Ctrl + Alt + F5** refreshes all). Then undo the change and refresh again.

## Step 6 – Create visualization

1. Click inside **PT_Region** > **PivotTable Analyze > PivotChart > Clustered Column > OK**.
1. Create a **Line** PivotChart for **PT_Month**, a **Bar** chart for **PT_Product (Top 5)** and a **Column** chart for **PT_Category** (Profit).
1. Move all four charts to a new sheet **Analysis** and align them with the slicers and timeline on top.
1. Hide the field buttons: right-click a chart > **Hide All Field Buttons on Chart**.

    ![Interactive sales analysis report](Linked_image_Files/lab10_step6_analysis-sheet.png)

## Expected Output

| Question | Answer |
|---|---|
| 1. Total sales by region | South ₹1,92,86,136.00; West ₹1,71,15,952.50; North ₹1,46,06,054.50; East ₹1,40,88,106.50 |
| 2. Highest-revenue product | **Laptop**, ₹2,89,63,000 (44.5% of net sales) |
| 3. Best salesperson | **Rohan Mehta**, ₹99,84,909 (South) |
| 4. Monthly trend | Lowest Jan ₹41,63,797.50; highest **Dec ₹73,28,612** |
| 5. Most profitable category | **Electronics**, profit ₹80,49,275 (Furniture ₹44,76,525; Office Supplies ₹5,53,989.50) |
| Margin% by category | Electronics 17.1%, Furniture 27.8%, Office Supplies 31.1% |
| Grand total | ₹6,50,96,249.50 |

## Screenshot of final result

![Final interactive analysis sheet with slicers](Linked_image_Files/lab10_final_interactive.png)

## Business Insights

- Laptops account for 44.5% of revenue, so a stock-out on one product would affect almost half of revenue.
- Electronics bring the **most profit in rupees** but have the **lowest margin** (17.1%). Furniture and Office Supplies earn better margins, so upselling these categories raises profitability.
- Rohan Mehta (South) tops the ranking; South is also the best region, so the team is a candidate for sharing best practice.

## Practice Questions

1. What are the four areas of a PivotTable field list?
1. Why can a calculated field `Profit/NetSales` be better than averaging margin percentages?
1. How do you make one slicer filter several PivotTables?
1. Why does a PivotTable not update automatically when source data changes?
1. How do you show the top 5 products only?

<details>
<summary>Answer key</summary>

1. Filters, Columns, Rows, Values.
1. It divides total profit by total sales, a weighted result.
1. Right-click the slicer > Report Connections.
1. Pivots keep a cache; you must Refresh.
1. Row Labels filter > Value Filters > Top 10 > 5.

</details>

## Lab Assignment

Create a pivot of **Region (Rows) × Category (Columns)** with NetSales as values and **Show Values As: % of Row Total**. Add a slicer for **Salesperson**. Write three recommendations for the Sales Director. Optional: use the **Walmart Sales** dataset linked in the dataset guide and build the same pivots by store.
