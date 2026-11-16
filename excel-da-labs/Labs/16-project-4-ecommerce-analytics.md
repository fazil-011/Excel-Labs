---
lab:
  title: Project 4 - E-commerce Analytics
  module: Module 13 - Real-World Projects
  description: In this capstone project, you analyse order lines, segment customers by spend, evaluate product and country performance, study sales trends and build a revenue dashboard.
  duration: 150 minutes
  level: 300
  islab: true
  primarytopics:
    - Excel
---

# Project 4 – E-commerce Analytics

**Lab Number:** 16
**Lab Title:** E-commerce Analytics – Orders, Customers and Revenue
**Module:** Module 13 – Real-World Projects
**Difficulty Level:** Advanced
**Estimated Time:** 150 minutes
**Prerequisites:** Labs 04, 05, 10, 11, 12.
**Excel version note:** `UNIQUE` and dynamic arrays need Excel 2021 or Microsoft 365. This lab uses a PivotTable for the customer list so it also works in Excel 2016.

## Problem Statement

An online gift shop has 1,510 order lines but no view of customer value, popular products or country performance. Marketing cannot target its best customers.

## Business Scenario

The marketing head asks: *What is our revenue and average order value? Which products and countries perform best? Who are our best customers? How do sales change month to month?*

## Learning Objectives

- Calculate revenue and average order value from order lines.
- Segment customers into High, Medium and Low value with percentiles.
- Analyse product, country and monthly performance.
- Build an e-commerce revenue dashboard.

## Dataset Download

| Dataset | Link |
|---|---|
| **Practice dataset (provided, synthetic):** ecommerce_orders.csv, 1,510 rows | [Download](../Allfiles/Datasets/ecommerce_orders.csv) |
| **Real-world option:** E-commerce Data (Kaggle) | https://www.kaggle.com/datasets/carrie1/ecommerce-data |

> *The Kaggle dataset is a UK online retailer's transactions (it is much larger, so Power Query is recommended instead of opening it directly, and it needs cleaning such as cancellations and missing customer IDs). It is not included in this package. The provided file uses the same column names so that the formulas also work on the real data.*

## Dataset Column Description (Data dictionary)

| Column | Type | Description |
|---|---|---|
| InvoiceNo | Whole number | Order number; one order has several lines |
| InvoiceDate | Date | Order date in 2024 |
| CustomerID | Whole number | Customer identifier |
| Country | Text | United Kingdom, Germany, France, India, USA |
| StockCode | Text | Product code |
| Description | Text | Product name |
| Quantity | Whole number | Units on the line |
| UnitPrice | Number | Price per unit |

## Required Excel functions

`SUMIFS`, `COUNTIFS`, `PERCENTILE.INC`, `IF`, `IFS`, `XLOOKUP` or `INDEX/MATCH`, `ROUND`, `TEXT`, PivotTables (distinct count), slicers, timelines.

## Step 1 – Import the dataset

1. Open **ecommerce_orders.csv**, save as **Ecommerce-Analytics.xlsx**, convert to table **tblOrders**, rename the sheet **Data**.

## Step 2 – Inspect the data

1. Rows: **1,510**; no blanks.
1. Orders: distinct **InvoiceNo = 500** (use **Remove Duplicates** on a copy of the column and count, or a pivot with **Add to Data Model > Distinct Count**).
1. Customers: distinct **CustomerID = 145**.
1. Check for negative or zero quantity (**Filter > Number Filters**). None in this file; in the Kaggle file, negative quantities are cancellations and must be handled.

## Step 3 – Apply the Excel operation

Add columns:

| Column | Formula |
|---|---|
| LineTotal | `=[@Quantity]*[@UnitPrice]` |
| Month | `=MONTH([@InvoiceDate])` |
| MonthName | `=TEXT([@InvoiceDate],"mmm")` |

![Order lines with LineTotal](Linked_image_Files/p4_step3_line-total.png)

## Step 4 – Apply formulas

| KPI | Formula | Expected |
|---|---|---|
| Revenue | `=SUM(tblOrders[LineTotal])` | 68,372.07 |
| Orders | distinct InvoiceNo | 500 |
| Customers | distinct CustomerID | 145 |
| Average Order Value | `=Revenue/Orders` | 136.74 |
| Units sold | `=SUM(tblOrders[Quantity])` | 18,541 |

### Customer segmentation

1. Create a pivot (sheet **Customers**): **CustomerID** in Rows, **Sum of LineTotal** in Values. Copy it to a normal range starting at **A1** with headers **CustomerID, Spend**.
1. In **E1** type **P25**, in **F1** type **P75**. Calculate the thresholds below them:

    | Cell | Formula | Expected |
    |---|---|---|
    | E2 | `=PERCENTILE.INC($B$2:$B$146,0.25)` | 219.20 |
    | F2 | `=PERCENTILE.INC($B$2:$B$146,0.75)` | 626.53 |

1. In **C1** type **Segment** and in **C2** enter:

    ```excel
    =IF(B2>=$F$2,"High",IF(B2<=$E$2,"Low","Medium"))
    ```

    Fill it down to **C146**.

1. Count segments with `=COUNTIF($C$2:$C$146,"High")` and so on.

> *Percentile-based segments adapt to the data. A fixed threshold such as "spend over 500" is simple but needs updating when prices change.*

## Step 5 – Analyse the results

| Question | Pivot / formula | Answer |
|---|---|---|
| Revenue by country | Country × LineTotal | United Kingdom 34,534.84 (50.5%); USA 8,211.63; India 9,189.12; France 9,185.97; Germany 7,250.51 |
| Top product by revenue | Description × LineTotal | **Regency Cakestand 3 Tier** 18,576.75 |
| Top product by units | Description × Quantity | **Metal Lantern** 1,784 units (Set of 3 Cake Tins 1,749) |
| Monthly revenue | Month × LineTotal | Peak **April 7,693.40**; lowest **July 3,904.29** |
| Customer segments | Segment counts | High 37, Medium 71, Low 37 |
| Top customers | Sort Spend descending | 10023 (1,544.33), 10087 (1,521.81), 10016 (1,397.88) |

Monthly revenue: Jan 6,821.61; Feb 4,275.72; Mar 6,032.06; Apr 7,693.40; May 6,306.82; Jun 4,925.74; Jul 3,904.29; Aug 4,925.87; Sep 7,149.85; Oct 5,952.80; Nov 5,359.43; Dec 5,024.48.

![Segmentation table and pivots](Linked_image_Files/p4_step5_segments.png)

## Step 6 – Create visualization

Build the **E-commerce Revenue Dashboard**:

1. KPI cards: Revenue, Orders, Customers, AOV, Units.
1. Charts: monthly revenue line, country bar, top-5 product column, segment doughnut (High/Medium/Low).
1. Slicers: Country, Month; Timeline for InvoiceDate.
1. A small **Top 10 customers** table sorted by spend with data bars.

    ![E-commerce dashboard](Linked_image_Files/p4_step6_dashboard.png)

## Expected Output

Dashboard with revenue ₹/£/$ **68,372.07** (use the currency symbol that suits your dataset; the provided file is currency-neutral), 500 orders, 145 customers, AOV 136.74 and the charts above.

## Screenshot of final result

![Final e-commerce dashboard](Linked_image_Files/p4_final_dashboard.png)

## Business Insights

1. **The home market dominates:** the United Kingdom brings in 50.5% of revenue; the next four countries each contribute 11–13%.
1. **Customer value:** High-value customers (37, one quarter of customers) spend at least 626.53 each; protect them with loyalty benefits and early access.
1. **Seasonality:** April and September are strong; July is the weakest month. Plan promotions for the summer dip.
1. **Product focus:** the Regency Cakestand earns the most revenue from fewer units because of its higher price, while Metal Lantern sells the most units.
1. Recommendation: a retention campaign for the Low segment (37 customers) and cross-selling cakestands to Medium customers.

## Interview questions based on this project

1. What is the difference between an order and an order line, and how did you count orders?
1. How did you segment customers, and why use percentiles?
1. How would you treat cancellations and missing customer IDs in the real Kaggle data?
1. Which metric would you use to find loyal customers besides total spend?
1. How would you calculate repeat purchase rate in Excel?

## Practice Questions

1. How do you count distinct customers in a PivotTable?
1. Why is average order value not `AVERAGE(LineTotal)`?
1. What does `PERCENTILE.INC(range,0.75)` return?
1. How would you flag customers whose last order was more than 90 days before 31-Dec-2024?
1. Why is a pie chart not recommended for countries here?

<details>
<summary>Answer key</summary>

1. Add the data to the Data Model and use Distinct Count.
1. A line is not an order; AOV = revenue ÷ distinct orders.
1. The value below which 75% of the data lies.
1. Get each customer's last date with a pivot (Max of InvoiceDate) and use `=IF(DATE(2024,12,31)-lastdate>90,"At risk","Active")`.
1. Five countries with one dominant slice, and small slices are hard to compare; a sorted bar is clearer.

</details>

## Lab Assignment

Compute **RFM-lite**: Recency (days since last order to 31-Dec-2024), Frequency (distinct orders) and Monetary (spend) per customer with a pivot and formulas. Create a segment label such as *Champions* and *At Risk*. Submit the workbook and a one-page recommendation.
