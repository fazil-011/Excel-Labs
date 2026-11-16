---
lab:
  title: Advanced Excel dashboards - Sales Analytics Dashboard
  module: Module 12 - Advanced Excel Dashboards
  description: In this lab, you plan KPIs and build a complete interactive Sales Analytics Dashboard with KPI cards, pivot charts, slicers, a timeline, conditional formatting and a dynamic title.
  duration: 120 minutes
  level: 300
  islab: true
  primarytopics:
    - Excel
---

# Sales Analytics Dashboard

**Lab Number:** 12
**Lab Title:** Build a Complete Interactive Sales Analytics Dashboard
**Module:** Module 12 – Advanced Excel Dashboards
**Difficulty Level:** Advanced
**Estimated Time:** 120 minutes
**Prerequisites:** Labs 04, 09, 10 and 11.
**Excel version note:** Works in Excel 2016+ (slicers and timelines need 2013+). `GETPIVOTDATA` works in all versions. Dynamic arrays are not required.

## Problem Statement

Executives want a single page that shows how the business is doing and let them filter by region, category and period without calling the analyst.

## Business Scenario

Build a dashboard with: **Total Revenue, Total Profit, Total Orders, Average Order Value, Monthly Revenue Trend, Region-wise Sales, Product Performance, Category Analysis** and **interactive slicers**.

## Learning Objectives

- Plan a dashboard: audience, questions, KPIs, layout.
- Build KPI cards that react to slicers.
- Create connected pivot charts, slicers and a timeline.
- Apply professional formatting and protect the dashboard.

## Dataset Download

| File | Records | Link |
|---|---|---|
| sales_transactions.csv | 1,200 | [Download](../Allfiles/Datasets/sales_transactions.csv) |

## Dataset Column Description

See Lab 03.

## Step 1 – Import the dataset

1. Open **sales_transactions.csv**, save as **Sales-Dashboard.xlsx**, make table **tblSales** and add **NetSales**, **Profit**, **Month** (Lab 09).
1. Create **four sheets**: **Dashboard**, **Calc**, **Pivots**, **Data**. Move the table to **Data**.

### Plan before you build

| Question | Answer for this dashboard |
|---|---|
| Audience | Sales Director and regional managers |
| Decisions | Where to invest, which products to push, which month to plan for |
| KPIs | Revenue, Profit, Orders, Average Order Value, Margin % |
| Filters | Region, Category, Date range |
| Layout | KPIs on top, trend in the middle, comparisons at the bottom, slicers on the left |

    ┌──────────────────────────────────────────────────────────┐
    │ SALES ANALYTICS DASHBOARD – 2024 (dynamic title)         │
    ├────────┬─────────┬─────────┬─────────┬─────────┬─────────┤
    │ Slicers│ Revenue │ Profit  │ Orders  │  AOV    │ Margin% │
    │ Region ├─────────┴─────────┴─────────┴─────────┴─────────┤
    │Category│        Monthly Revenue Trend (line)             │
    │Timeline├───────────────────┬─────────────────────────────┤
    │        │ Region Sales (col)│ Category Share (doughnut)   │
    │        ├───────────────────┴─────────────────────────────┤
    │        │      Top Products (bar)                         │
    └────────┴─────────────────────────────────────────────────┘

## Step 2 – Inspect the data

1. Check totals: Net Sales **6,50,96,249.50**, Profit **1,30,79,789.50**, Orders **1,200**.
1. Add a column **OrderCount** with value `1` to count orders in pivots (optional) or use **Count of OrderID**.

## Step 3 – Apply the Excel operation

### Build the pivot tables (sheet Pivots)

Create all five from the **same** table and place them side by side. Name them with **PivotTable Analyze > PivotTable Name**.

| Name | Rows | Columns | Values |
|---|---|---|---|
| PT_KPI | (none) | (none) | Sum of NetSales, Sum of Profit, Count of OrderID |
| PT_Month | OrderDate grouped by Months | – | Sum of NetSales |
| PT_Region | Region | – | Sum of NetSales |
| PT_Category | Category | – | Sum of NetSales |
| PT_Product | Product (Top 8 by NetSales) | – | Sum of NetSales |

> *Create the first pivot, **copy and paste it** to make others from the same cache, so a slicer can control all of them.*

## Step 4 – Apply formulas

On the **Calc** sheet build the KPI values with `GETPIVOTDATA`, which reads from the pivot and responds to slicers:

| Cell | Formula |
|---|---|
| B2 Revenue | `=GETPIVOTDATA("NetSales",Pivots!$A$3)` |
| B3 Profit | `=GETPIVOTDATA("Profit",Pivots!$A$3)` |
| B4 Orders | `=GETPIVOTDATA("OrderID",Pivots!$A$3)` |
| B5 AOV | `=IFERROR(B2/B4,0)` |
| B6 Margin % | `=IFERROR(B3/B2,0)` |

(Adjust `Pivots!$A$3` to the top-left cell of **PT_KPI**. To get the correct syntax automatically, type `=` in a cell and click a value in the pivot.)

Dynamic title in **Calc!B8**:

```excel
="Sales Analytics Dashboard – 2024"
```

Extend it later with slicer text, for example:

```excel
="Sales Analytics Dashboard – "&IF(COUNTA(Pivots!$H$3:$H$6)=1,Pivots!$H$3,"All Regions")
```

> *Dynamic titles need a pivot (or helper range) that shows the selected slicer item. Keep it simple: a static title is acceptable for the first build.*

## Step 5 – Analyse the results

### KPI cards

1. On **Dashboard**, turn off gridlines (**View > Gridlines**) and set column widths to 2 for a grid look (optional).
1. Merge-less cards: select **B3:C5**, set a dark blue fill, and use **Home > Alignment > Center Across Selection**.
1. In **B3** type `=Calc!B2/10000000` (Revenue in crores) and in **B4** the label *Total Revenue*. Format **B3** with the custom format `[$₹-4009] 0.00" Cr"`; the value ₹6,50,96,249.50 shows as **₹6.51 Cr**.

    > *Custom number formats do not change the value; they only change how it appears.*

1. Repeat for Profit, Orders (format `#,##0`), AOV (`₹#,##0`), Margin% (`0.0%`).

    ![KPI cards design](Linked_image_Files/lab12_step5_kpi-cards.png)

### Charts

1. From each pivot create a **PivotChart**: **PT_Month → Line**, **PT_Region → Clustered Column**, **PT_Category → Doughnut**, **PT_Product → Bar**.
1. Move each chart to **Dashboard** (Cut/Paste), resize to fit the layout grid and remove field buttons.
1. Formatting: one font (e.g. Segoe UI), blue palette, remove gridlines, chart titles of 12 pt.

### Slicers and timeline

1. Click **PT_KPI** > **Insert Slicer** > **Region**, **Category**. Insert a **Timeline** for **OrderDate**.
1. Right-click each slicer > **Report Connections** > tick **all five pivots**.
1. Move slicers and timeline to the left panel on **Dashboard**. Set **Slicer > Columns = 1** (Region: 4 buttons).

## Step 6 – Create visualization

### Conditional formatting and polish

1. Add a small table under the region chart: Region, Net Sales, Margin %. Use **Conditional Formatting > Data Bars** for Net Sales and **Icon Sets** (3 arrows) for Margin %.
1. **Protect** the dashboard: **Review > Protect Sheet** (allow selecting unlocked cells; slicers are still usable if you tick *Use PivotTable & PivotChart* and unlock slicers via **Slicer > Properties > Locked** off).
1. Hide the **Calc**, **Pivots** and **Data** sheets (right-click > **Hide**) before sharing.
1. Set **View > Zoom** so that the dashboard fits on one screen at 100%.

    ![Final dashboard](Linked_image_Files/lab12_step6_dashboard-final.png)

### Usability checklist

| Principle | Check |
|---|---|
| One screen | No scrolling |
| Clear hierarchy | KPIs first, details below |
| Consistent colour | One palette, one highlight colour |
| Labels | Units and period visible |
| Interactivity | Slicers tested; reset button (Clear Filter) |
| Reporting best practice | Data source and refresh date noted at the bottom |

## Expected Output

| Dashboard element | Expected (no filters) |
|---|---|
| Total Revenue | ₹6,50,96,249.50 (card shows ₹6.51 Cr) |
| Total Profit | ₹1,30,79,789.50 |
| Total Orders | 1,200 |
| Average Order Value | ₹54,246.87 |
| Margin % | 20.1% |
| Region chart | South highest |
| Category doughnut | Electronics 72.5% |
| Product bar | Laptop first |
| After selecting Region = West | Revenue ₹1,71,15,952.50, Orders 307, Profit ₹34,06,322.50 |
| After selecting Category = Furniture | Revenue ₹1,61,10,625, Profit ₹44,76,525 |

## Screenshot of final result

![Final Sales Analytics Dashboard](Linked_image_Files/lab12_final_dashboard.png)

## Business Insights

- The business earns about ₹20 on every ₹100 of net sales, but Electronics earn only 17%; a product-mix shift is the cheapest way to improve margin.
- West and South together contribute 56% of revenue.
- A dashboard that responds to slicers replaces a monthly Excel report that previously took hours.

## Practice Questions

1. Why use `GETPIVOTDATA` for KPI cards instead of linking to a normal cell?
1. How do you make one slicer control several pivots?
1. Why should you hide the helper sheets?
1. What is a custom number format and why use one for KPIs?
1. Name three dashboard design principles.

<details>
<summary>Answer key</summary>

1. It follows the pivot when filters change; a plain cell reference can point at the wrong cell.
1. Right-click slicer > Report Connections.
1. To keep users from editing calculations and keep the dashboard clean.
1. It changes how a value appears (for example crores) without changing the value.
1. One screen, clear hierarchy, consistent colours, labelled units, interactivity.

</details>

## Lab Assignment

Add a seventh KPI **Profit per Order** and a **top 5 salespeople** chart. Add a *Reset* instruction text box. Export the dashboard as PDF (**File > Export > PDF**) and write a half-page note describing the story the dashboard tells.
