---
lab:
  title: Project 1 - Retail Sales Analytics
  module: Module 13 - Real-World Projects
  description: In this capstone project, you analyse retail sales transactions, identify best-selling products, calculate revenue and profit, analyse regional performance and discount impact, and build an interactive dashboard.
  duration: 150 minutes
  level: 300
  islab: true
  primarytopics:
    - Excel
---

# Project 1 – Retail Sales Analytics

**Lab Number:** 13
**Lab Title:** Retail Sales Analytics
**Module:** Module 13 – Real-World Projects
**Difficulty Level:** Advanced
**Estimated Time:** 150 minutes
**Prerequisites:** Labs 03, 04, 09, 10, 12.

## Problem Statement

A retail company with four regions and eight salespeople sold ₹6.5 crore of products in 2024. Management does not know which products and regions create profit, whether discounts hurt margin, or where to focus in 2025.

## Business Scenario

You are the retail data analyst. The CEO asks for a decision-ready dashboard that answers: *What sells most? Where do we make money? Do discounts pay off? Who and what should we focus on?*

## Learning Objectives

Business objectives of the project:

- Analyse sales transactions end to end.
- Identify best-selling products by revenue **and** by units.
- Calculate revenue, profit and margin.
- Analyse regional and salesperson performance.
- Measure the impact of discounts on margin.
- Build an interactive dashboard.

## Dataset Download

| Dataset | Link |
|---|---|
| **Practice dataset (provided):** sales_transactions.csv, 1,200 rows | [Download](../Allfiles/Datasets/sales_transactions.csv) |
| **Real-world option:** Superstore Sales Dataset (Kaggle) | https://www.kaggle.com/datasets/vivek468/superstore-dataset-final |

> *The Kaggle file must be downloaded by you from Kaggle (login required) and is **not** included in this package. Its columns differ from the practice dataset (for example it has Sales and Profit already calculated). If you use it, map the columns as shown below. If characters look wrong when importing, set **File Origin** to **1252 Western European** in the Power Query preview.*

## Dataset Column Description (Data dictionary)

| Column | Type | Description |
|---|---|---|
| OrderID | Text | Unique order line ID |
| OrderDate | Date | Order date in 2024 |
| Region | Text | North / South / East / West |
| Salesperson | Text | Responsible salesperson |
| Category | Text | Electronics / Furniture / Office Supplies |
| Product | Text | 12 products |
| Quantity | Whole number | Units sold |
| UnitPrice | Currency | Selling price per unit before discount |
| UnitCost | Currency | Cost per unit |
| DiscountPct | Number | Discount % (0, 5, 10, 15) |

Kaggle Superstore mapping (if you choose it): Order Date → OrderDate, Region → Region, Category → Category, Product Name → Product, Quantity → Quantity, Discount (0–1) → DiscountPct ÷ 100, Sales → NetSales, Profit → Profit.

## Required Excel functions

`SUM`, `COUNTA`, `AVERAGE`, `IF`, `IFS`, `SUMIFS`, `COUNTIFS`, `INDEX`, `MATCH`, `LARGE`, `RANK`, `ROUND`, `GETPIVOTDATA`, PivotTables, Slicers, Timelines, PivotCharts, Conditional Formatting, Power Query (optional).

## Step 1 – Import the dataset

1. Open **sales_transactions.csv** and save as **Retail-Analytics.xlsx**.
1. Convert to table **tblSales** and rename the sheet **Data**.
1. (Optional) Load the data through Power Query (**Data > From Text/CSV > Transform Data**) so that the project is refreshable.

    ![Importing the data](Linked_image_Files/p1_step1_import.png)

## Step 2 – Inspect the data

1. Check: 1,200 rows, no blanks, dates between 01-Jan-2024 and 31-Dec-2024.
1. Check that each Region has 2 salespeople: pivot **Region × Salesperson** (count).
1. Check **DiscountPct** has four values: 0, 5, 10, 15.

## Step 3 – Apply the Excel operation

Add calculated columns:

| Column | Formula |
|---|---|
| NetSales | `=[@Quantity]*[@UnitPrice]*(1-[@DiscountPct]/100)` |
| TotalCost | `=[@Quantity]*[@UnitCost]` |
| Profit | `=[@NetSales]-[@TotalCost]` |
| Month | `=MONTH([@OrderDate])` |
| MonthName | `=TEXT([@OrderDate],"mmm")` |
| DiscountBand | `=IFS([@DiscountPct]=0,"No discount",[@DiscountPct]<=5,"5%",[@DiscountPct]<=10,"10%",TRUE,"15%")` |

(Excel 2016 without `IFS`: use nested `IF`.)

![Calculated columns](Linked_image_Files/p1_step3_calculated-columns.png)

## Step 4 – Apply formulas

On a **Calc** sheet build the headline KPIs:

| KPI | Formula | Expected |
|---|---|---|
| Revenue | `=SUM(tblSales[NetSales])` | ₹6,50,96,249.50 |
| Profit | `=SUM(tblSales[Profit])` | ₹1,30,79,789.50 |
| Margin | `=B3/B2` | 20.1% |
| Orders | `=COUNTA(tblSales[OrderID])` | 1,200 |
| Units | `=SUM(tblSales[Quantity])` | 6,589 |

Best-seller formulas:

```excel
=INDEX(Pivots!A4:A15,MATCH(MAX(Pivots!B4:B15),Pivots!B4:B15,0))
```

Rank products: `=RANK(B4,$B$4:$B$15,0)`.

## Step 5 – Analyse the results

Create pivots (sheet **Pivots**) and answer:

| Business question | Pivot setup | Answer |
|---|---|---|
| Top product by revenue | Product × NetSales, sorted | **Laptop** ₹2,89,63,000 (44.5%) |
| Top product by units | Product × Sum Quantity | **Headphones** 613 units; then Notebook Pack 603 and Stapler 581 |
| Revenue by region | Region × NetSales | South ₹1.93 Cr, West ₹1.71 Cr, North ₹1.46 Cr, East ₹1.41 Cr |
| Profit by salesperson | Salesperson × Profit | **Rohan Mehta** ₹20,31,519; Vikram Singh lowest ₹12,28,560 |
| Margin by category | Category × (Profit ÷ NetSales) calculated field | Electronics 17.1%, Furniture 27.8%, Office Supplies 31.1% |
| Peak month | Month × NetSales | December, ₹73,28,612 (profit ₹15,78,242) |
| **Discount impact** | DiscountBand × (Profit ÷ NetSales), and count of orders | No discount 24.2% (614 orders); 5% → 20.8% (212); 10% → 15.4% (195); 15% → 10.4% (179) |

![Pivot tables for the project](Linked_image_Files/p1_step5_pivots.png)

## Step 6 – Create visualization

Build the **Retail Executive Dashboard** (reuse the layout of Lab 12):

1. Top row: Revenue, Profit, Margin %, Orders, Units cards.
1. Charts: monthly revenue line, region column, category doughnut, top-6 products bar, discount band vs margin column.
1. Slicers: Region, Category, Salesperson; Timeline: OrderDate.
1. Conditional formatting: red text for months where margin is below 18%.

    ![Retail Executive Dashboard](Linked_image_Files/p1_step6_dashboard.png)

## Expected Output

A one-page dashboard that shows the KPIs above, responds to slicers, and has the insights panel below.

## Screenshot of final result

![Final Retail dashboard](Linked_image_Files/p1_final_dashboard.png)

## Business Insights

1. **Concentration risk:** Laptops are 44.5% of revenue and Electronics 72.5%. Protect supply and pricing there.
1. **Discounts destroy margin:** margin falls from 24.2% with no discount to 10.4% at 15%. 586 of 1,200 orders (49%) carry a discount. Review discount approval rules.
1. **Mix opportunity:** Furniture and Office Supplies have higher margin than Electronics; bundle them with laptops.
1. **Regional focus:** South leads in volume; East has the fewest orders (274) and the lowest sales: target East for growth.
1. **Seasonality:** December peak; plan inventory and campaigns for Q4 and run promotions in January.

## Interview questions based on this project

1. How did you calculate profit and margin, and why do you divide total profit by total sales instead of averaging margins?
1. Which function or feature did you use to find the top product, and what alternatives exist?
1. How would you prove that discounts reduce margin?
1. How did you make the KPI cards respond to slicers?
1. If data for 2025 arrives monthly, how would you automate the report?

## Practice Questions

1. Which pivot setting shows each category's share of the total?
1. Why is the best-selling product by units different from the best by revenue?
1. How do you create a discount band without nested IF in Excel 365?
1. What is the risk of a dashboard with more than 10 charts?
1. How can Power Query help in this project?

<details>
<summary>Answer key</summary>

1. Show Values As > % of Grand Total.
1. Unit prices differ widely (laptop ₹55,000 vs pen ₹180).
1. `IFS`, `LOOKUP` or a `XLOOKUP` against a band table.
1. Too much information hides the message.
1. It combines new monthly files and refreshes with one click.

</details>

## Lab Assignment

Write a **one-page executive summary** (three insights, three recommendations, one risk) and export the dashboard to PDF. Optional: repeat the analysis with the Kaggle Superstore dataset and compare which categories lose money.
