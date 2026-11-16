---
lab:
  title: Data cleaning - Clean an unstructured customer sales dataset
  module: Module 7 - Data Cleaning
  description: In this lab, you clean a deliberately messy dataset with duplicates, blanks, inconsistent names and cities, invalid dates and text amounts, and document before-and-after results.
  duration: 90 minutes
  level: 200
  islab: true
  primarytopics:
    - Excel
---

# Clean an Unstructured Customer Sales Dataset

**Lab Number:** 07
**Lab Title:** Clean an Unstructured Customer Sales Dataset
**Module:** Module 7 – Data Cleaning
**Difficulty Level:** Intermediate
**Estimated Time:** 90 minutes
**Prerequisites:** Labs 01–06.

## Problem Statement

A customer sales export contains duplicate rows, missing values, inconsistent names and cities, invalid dates and amounts typed as text. Any report built on it would be wrong.

## Business Scenario

The finance team asked for revenue by city and by month, but totals do not match the billing system. Your job is to clean the file, document every change, and show before-and-after results.

## Learning Objectives

- Identify missing values, duplicates, inconsistent text, invalid dates and bad numbers.
- Use Remove Duplicates, Find and Replace, Go To Special, Text to Columns, Flash Fill, `TRIM`, `PROPER`, `IFERROR`, `DATEVALUE`, `VALUE`.
- Follow data cleaning best practices: never overwrite raw data and always log changes.

## Dataset Download

| File | Records | Link |
|---|---|---|
| customer_sales_messy.csv | 160 | [Download](../Allfiles/Datasets/customer_sales_messy.csv) |
| customer_sales_clean_reference.csv (answer key) | 150 | [Download](../Allfiles/Datasets/customer_sales_clean_reference.csv) |

> *Open the reference file only after you finish. Some issues (for example the text amount "five thousand") cannot be recovered; your cleaned file will differ from the reference in those rows, and that is correct: flag, do not guess.*

## Dataset Column Description

| Column | Description | Known problems |
|---|---|---|
| OrderID | M0001 – M0150 | 10 exact duplicate rows (160 rows in the file, 150 unique) |
| CustomerName | Customer name | 51 rows with extra spaces or wrong case (48 after duplicates are removed) |
| City | City | 13 blanks (12 after duplicates are removed); 19 distinct spellings of 6 real cities |
| OrderDate | Order date | 10 invalid values in the file (`N/A` ×5, `31/02/2024` ×4, `2024-13-05` ×1); 9 after duplicates are removed; some valid dates are in `dd-mmm-yyyy` form |
| Category | Product category | Clean |
| Quantity | Units | 3 blanks |
| Amount | Order amount | 15 text values in the file, 14 after duplicates are removed (`Rs. 25333`, `$19629`, `15277 INR`, `five thousand`) |
| Email, Phone | Contact data | Clean |

## Step 1 – Import the dataset

1. Open **customer_sales_messy.csv** and save as **Cleaning-Lab.xlsx**.
1. Rename the sheet **Raw** and **never edit it again**. Right-click the tab > **Move or Copy > Create a copy** and name the copy **Working**.
1. Create a third sheet **Log** with headers **Issue**, **Rows affected**, **Action**, **Result**.

## Step 2 – Inspect the data (before-cleaning results)

On **Raw**, record these in the Log:

| Check | How |
|---|---|
| Total rows | Select column A: **Count = 160** |
| Duplicate rows | **Home > Conditional Formatting > Highlight Cells Rules > Duplicate Values** on A:A (10 extra) |
| Blank cells | Select **A1:I161**, **F5 > Special > Blanks** (16 blank cells: 13 City + 3 Quantity) |
| Distinct cities | Apply a filter to **City** and count items (20 entries including a blank) |
| Text numbers | In a spare cell, `=SUMPRODUCT(--ISTEXT(G2:G161))` returns **15** text amounts |
| Mixed dates | Filter **OrderDate**: valid dates group together; text values appear separately |

Also compute the **before total**: `=SUM(G2:G161)`. It returns **₹33,11,046**; it ignores text and counts duplicate rows, so it is wrong in both directions.

![Before-cleaning inspection](Linked_image_Files/lab07_step2_before.png)

## Step 3 – Apply the Excel operation

Work on the **Working** sheet in this order.

### Part A – Remove duplicates

1. Click any cell in the data and select **Data > Remove Duplicates**.
1. Tick **all** columns (so only exact duplicates are removed) and select **OK**.
1. Excel reports *"10 duplicate values found and removed; 150 unique values remain."* Log it.

### Part B – Standardise names

1. Insert a column next to CustomerName: header **CleanName**. In **C2** type `=PROPER(TRIM(B2))`, fill down, **Copy > Paste Special > Values**, then delete the old **CustomerName** column.
1. **Flash Fill option:** type the first name from the first row in a new column, press **Enter**, then **Ctrl + E**. Excel guesses the pattern. Always check the result.
1. **Text to Columns option:** select the clean name column > **Data > Text to Columns > Delimited > Space > Finish** to split into first and last name (insert two blank columns to the right first so you do not overwrite data).

### Part C – Standardise cities with Find & Replace

1. Select the **City** column and press **Ctrl + H**.
1. Replace each variant (keep **Match case** off, **Match entire cell contents** on):

    | Find | Replace with |
    |---|---|
    | Bombay | Mumbai |
    | New Delhi | Delhi |
    | Hyd | Hyderabad |

1. Add a helper column: `=PROPER(TRIM(D2))` then paste values over the City column. This fixes `PUNE`, `delhi`, `MUMBAI ` and similar.
1. Verify: the filter list for **City** now shows **6** cities: Bengaluru, Chennai, Delhi, Hyderabad, Mumbai, Pune (plus blanks).

### Part D – Handle missing values

1. Select the **City** column, **F5 > Special > Blanks > OK**. Type **Unknown** and press **Ctrl + Enter**. All 12 remaining blanks fill at once.
1. For **Quantity** blanks (3 rows): do **not** guess silently. Add a column **QtyFlag** `=IF(F2="","Missing","OK")`. Then fill the blank with the **median** quantity: `=IF(F2="",MEDIAN($F$2:$F$151),F2)`. Log the assumption.

### Part E – Fix invalid dates

1. Add column **DateClean**:

    ```excel
    =IF(ISNUMBER(E2),E2,IFERROR(DATEVALUE(E2),""))
    ```

    > *If the cell already holds a real date (number), keep it; if it holds text that looks like a date, convert it; otherwise return an empty text so it can be reviewed.*

1. Format **DateClean** as `dd-mmm-yyyy`. Add **DateFlag**: `=IF(DateClean="","Invalid","OK")`. Filter on **Invalid** and read the count: **9**. Do not invent dates; send these rows back to the source owner.

### Part F – Convert amounts to numbers

```excel
=IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H2,"Rs. ",""),"INR",""),"$","")," ","")),"")
```

Rows such as `five thousand` return empty and are flagged **AmountFlag = Review**. After duplicates are removed there are **14** text amounts; **10** convert successfully and **4** need review.

### Part G – Finish

1. Convert formulas to values, delete helper columns, keep **OrderID, CleanName, City, DateClean, Category, Quantity, AmountClean**.
1. Format as a table **tblClean**.

## Step 4 – Apply formulas (quality checks)

| Check | Formula | Expected |
|---|---|---|
| Rows | `=COUNTA(tblClean[OrderID])` | 150 |
| Unique IDs | `=SUMPRODUCT(1/COUNTIF(tblClean[OrderID],tblClean[OrderID]))` | 150 |
| Blank cities | `=COUNTBLANK(tblClean[City])` | 0 |
| Distinct cities | Count in a filter | 6 real + Unknown |
| Invalid dates | `=COUNTIF(tblClean[DateFlag],"Invalid")` | 9 |
| Total amount | `=SUM(tblClean[AmountClean])` | **₹33,96,968** (4 unrecoverable amounts are excluded) |

## Step 5 – Analyse the results

1. Fill the **Log** sheet, one row per issue, with before and after counts.
1. Calculate the **impact**: *Raw revenue (text ignored, duplicates included)* vs *Clean revenue (duplicates removed, convertible text recovered)*.
1. Open **customer_sales_clean_reference.csv**; its total amount is **₹35,29,016**. Your total should be **₹33,96,968**, which is **₹1,32,048 lower** because of the 4 amounts typed as `five thousand`, which cannot be recovered. Explain the gap in the Log.

## Step 6 – Create visualization

1. Create a **before-and-after** table on the Log sheet: rows, distinct cities, blank cells, text amounts, invalid dates.
1. Insert a **Clustered Bar** chart that compares **Before** and **After** counts for duplicates, blanks, city variants and text amounts.

    ![Before and after cleaning comparison](Linked_image_Files/lab07_step6_before-after-chart.png)

## Expected Output

| Metric | Before | After |
|---|---|---|
| Rows | 160 | 150 |
| Duplicate rows | 10 | 0 |
| City entries (distinct) | 20 | 6 + "Unknown" |
| Blank cities | 13 | 0 (Unknown) |
| Blank quantities | 3 | 0 (median) |
| Names with spacing/case errors | 51 | 0 |
| Invalid dates | 10 | 9 flagged (not guessed; 1 was in a removed duplicate) |
| Text amounts | 15 | 4 flagged for review (10 recovered) |
| Total amount | ₹33,11,046 | ₹33,96,968 |

## Screenshot of final result

![Final cleaned dataset and log](Linked_image_Files/lab07_final_clean-dataset.png)

## Business Insights

- Before cleaning, revenue by city had 20 city labels, so Mumbai sales were split across `Mumbai`, `Bombay`, `mumbai` and `MUMBAI `. After cleaning, six cities give true totals.
- Duplicate rows inflated order counts by 6.7% (10 of 150).
- Documenting issues turns cleaning into a repeatable process and shows the business where the source system needs validation (see Lab 08).

## Practice Questions

1. Why must you keep an untouched copy of the raw data?
1. What does **Go To Special > Blanks** plus **Ctrl + Enter** do?
1. What is the difference between **Flash Fill** and a formula?
1. Why not fill an invalid date with today's date?
1. Why does `SUM` understate a column that contains `Rs. 5000` as text?

<details>
<summary>Answer key</summary>

1. To audit and redo cleaning if a mistake is made.
1. Selects all blank cells and fills them with the same typed entry.
1. Flash Fill gives static values based on a pattern; a formula stays dynamic.
1. It invents data and hides the problem.
1. SUM ignores text values.

</details>

## Lab Assignment

Create a **Data Quality Scorecard**: completeness % (non-blank cells ÷ total), uniqueness % (unique IDs ÷ rows), validity % (valid dates ÷ rows) for before and after. Submit the workbook and a 5-line cleaning report.
