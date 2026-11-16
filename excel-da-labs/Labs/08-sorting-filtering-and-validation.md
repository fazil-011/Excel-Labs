---
lab:
  title: Sorting, filtering and validation - Inventory Management worksheet
  module: Module 8 - Sorting, Filtering and Data Validation
  description: In this lab, you sort and filter an inventory list, use Advanced Filter, filter by colour, remove duplicates and build a stock-issue form with data validation dropdowns, input messages and error alerts.
  duration: 75 minutes
  level: 200
  islab: true
  primarytopics:
    - Excel
---

# Inventory Management Worksheet with Dropdown Product Selection

**Lab Number:** 08
**Lab Title:** Create an Inventory Management Worksheet
**Module:** Module 8 – Sorting, Filtering and Data Validation
**Difficulty Level:** Intermediate
**Estimated Time:** 75 minutes
**Prerequisites:** Labs 01–07.

## Problem Statement

The stores team tracks 36 products in a sheet. They cannot quickly find low-stock items, staff mistype product names, and nothing stops someone from issuing more stock than exists.

## Business Scenario

The store manager wants: (1) products sorted by category and stock, (2) an instant list of items to reorder, (3) a form where staff pick a product from a dropdown and enter the quantity issued, with errors blocked.

## Learning Objectives

- Sort by one level, multiple levels and a custom order.
- Use AutoFilter, text, number and date filters, filter by colour and Advanced Filter.
- Create dropdown lists, input messages and error alerts with Data Validation.
- Remove duplicates safely.

## Dataset Download

| File | Records | Link |
|---|---|---|
| inventory.csv | 36 | [Download](../Allfiles/Datasets/inventory.csv) |

## Dataset Column Description

| Column | Description |
|---|---|
| ProductID | P101 – P136 |
| ProductName | Product name |
| Category | Electronics, Furniture, Office Supplies |
| Supplier | TechSource Pvt Ltd, OfficeMart, FurniWorld, BrightSupplies |
| StockQty | Units on hand |
| ReorderLevel | Reorder when stock is at or below this level |
| UnitPrice | Price per unit, INR |

## Step 1 – Import the dataset

1. Open **inventory.csv**, save as **Inventory-Management.xlsx** and convert the data to a table named **tblInventory**.
1. Rename the sheet **Inventory**. Add sheets **Issue Form** and **Lists**.

## Step 2 – Inspect the data

1. Count rows: **36**.
1. Add a column **Status**: `=IF([@StockQty]<=[@ReorderLevel],"Reorder","OK")`.
1. Add **StockValue**: `=[@StockQty]*[@UnitPrice]`.
1. Apply conditional formatting to **Status**: **Highlight Cells Rules > Text that Contains > Reorder** with red fill. You use this colour in Part C.

## Step 3 – Apply the Excel operation

### Part A – Sorting

1. **Single level:** click **Category** header arrow > **Sort A to Z**.
1. **Multi-level:** **Data > Sort**. Level 1: **Category** A–Z. **Add Level**: **StockQty** Smallest to Largest. **OK**.
1. **Custom order:** **Data > Sort**, Level 1: **Category**, Order: **Custom List...** > type *Electronics, Furniture, Office Supplies* (one per line) > **Add > OK**.

### Part B – Filtering

1. Click the **Category** arrow, untick **Select All**, tick **Electronics**.
1. Click the **StockQty** arrow > **Number Filters > Less Than** > **50**.
1. Click the **ProductName** arrow > **Text Filters > Contains** > **Pro**.
1. Clear all filters (**Data > Clear**).

### Part C – Filter by colour

1. Click the **Status** arrow > **Filter by Color > Filter by Cell Color** > red. Four items are shown.
1. Clear the filter.

### Part D – Advanced Filter

1. On the **Lists** sheet, type a criteria range in **A1:B2**:

    | Category | StockQty |
    |---|---|
    | Office Supplies | <=20 |

1. Go to **Inventory**, **Data > Advanced**. Action: **Copy to another location**. List range: the table (`$A$1:$I$37`). Criteria range: `Lists!$A$1:$B$2`. Copy to: `Lists!$D$1`. **OK**.
1. Items with **Office Supplies** and stock ≤ 20 appear in D1 onward. Using two criteria on the **same row** means AND; criteria on **different rows** means OR.

### Part E – Remove duplicates (practice copy)

1. Copy the table to a new sheet **DupTest**. Copy rows 2–4 and paste below to create 3 duplicates.
1. **Data > Remove Duplicates**, tick all columns, **OK**. Excel reports 3 duplicates removed.

## Step 4 – Apply formulas

### Part A – Source lists and validation

1. On **Lists**, type the **Category** list in **G1:G3** (Electronics, Furniture, Office Supplies).
1. Go to **Inventory**, select a blank cell outside the table (for a test), then **Data > Data Validation > Allow: List > Source: `=Lists!$G$1:$G$3` > OK**. A dropdown appears.
1. Use the same method to build a **Supplier** list.

### Part B – Dropdown product selection form

On **Issue Form**:

| Cell | Content |
|---|---|
| A1 | Stock Issue Form |
| A3 | Select Product: |
| B3 | Dropdown (see below) |
| A4 | Current Stock: |
| B4 | `=IFERROR(INDEX(tblInventory[StockQty],MATCH($B$3,tblInventory[ProductName],0)),"")` |
| A5 | Reorder Level: |
| B5 | `=IFERROR(INDEX(tblInventory[ReorderLevel],MATCH($B$3,tblInventory[ProductName],0)),"")` |
| A6 | Quantity to Issue: |
| B6 | Input cell |
| A7 | Stock After Issue: |
| B7 | `=IF(AND(ISNUMBER(B4),ISNUMBER(B6)),B4-B6,"")` |
| A8 | Message: |
| B8 | `=IF(B7="","",IF(B7<=B5,"Reorder now","Stock OK"))` |

1. Select **B3** > **Data Validation > List**, Source: `=tblInventory[ProductName]` (if your version rejects table references, define the name **ProductList** with **Formulas > Name Manager** and use `=ProductList`).
1. **Input Message** tab: Title *Product*, Message *Choose a product from the list.*
1. **Error Alert** tab: Style **Stop**, Title *Invalid product*, Message *Please select a product from the dropdown.*
1. Select **B6** > **Data Validation > Allow: Whole number > between > Minimum 1 > Maximum `=B4`**. Input message: *Enter a quantity from 1 up to the current stock.* Error alert (Stop): *You cannot issue more than the stock on hand.*

    ![Data Validation dialog for B3](Linked_image_Files/lab08_step4_validation-dialog.png)

> *Error alert styles: **Stop** blocks the entry; **Warning** lets the user continue after confirming; **Information** only informs.*

## Step 5 – Analyse the results

1. Select **Laptop** in **B3**. It shows Stock **23**, Reorder level **30**. Enter **5** in **B6**: Stock After Issue **18** and message *Reorder now*.
1. Type **50** in **B6**. Excel blocks it with the error alert.
1. Type **Phone** in **B3**. Excel blocks it because it is not in the list.
1. Select **Data > Data Validation > Circle Invalid Data** on a column of existing data to find old entries that break the rule.

## Step 6 – Create visualization

1. Filter **Status = Reorder** and select **ProductName** and **StockQty** (and **ReorderLevel**).
1. Insert a **Clustered Column** chart with **StockQty** vs **ReorderLevel** for those four items.
1. Title it **Items Below Reorder Level**.

    ![Reorder chart](Linked_image_Files/lab08_step6_reorder-chart.png)

## Expected Output

| Check | Expected |
|---|---|
| Records | 36 |
| Items to reorder | 4: Laptop (23/30), Printer Paper Pro (9/10), Printer Paper Lite (17/50), Stapler Pro (7/20) |
| Total stock value `=SUM(tblInventory[StockValue])` | ₹3,33,88,660 |
| Laptop issue of 5 | After = 18, message "Reorder now" |
| Issue of 50 from Laptop | Blocked by error alert |

## Screenshot of final result

![Final Issue Form and filtered inventory](Linked_image_Files/lab08_final_issue-form.png)

## Business Insights

- Four products need immediate reordering; three are office supplies with low unit prices that are easy to overlook.
- Validation removes spelling errors and prevents negative stock, so the stock data stays reliable.

## Practice Questions

1. What is the difference between AutoFilter and Advanced Filter?
1. In an Advanced Filter, how do you write an OR condition?
1. Which Error Alert style blocks bad input?
1. How do you sort in your own order (for example Low, Medium, High)?
1. Why does Remove Duplicates need all columns ticked to remove only exact duplicates?

<details>
<summary>Answer key</summary>

1. Advanced Filter uses a separate criteria range, supports complex AND/OR logic and can copy results elsewhere.
1. Put the conditions on different rows of the criteria range.
1. Stop.
1. Sort > Order > Custom List.
1. Excel compares only the ticked columns; fewer ticks can remove different records.

</details>

## Lab Assignment

Add a **Supplier** dropdown and a **dependent Product dropdown** (hint: use `INDIRECT` with named ranges per category). Add a log table where each valid issue is recorded manually, then calculate remaining stock per product with `SUMIFS`. Submit the workbook.
