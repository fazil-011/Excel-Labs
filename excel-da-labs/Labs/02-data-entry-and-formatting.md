---
lab:
  title: Data entry and formatting - Employee Information Report
  module: Module 2 - Data Entry and Formatting
  description: In this lab, you turn a raw employee list into a professional, print-ready report using number and date formats, tables, conditional formatting, freeze panes and page layout.
  duration: 60 minutes
  level: 100
  islab: true
  primarytopics:
    - Excel
---

# Employee Information Report

**Lab Number:** 02
**Lab Title:** Create a Professional Employee Information Report
**Module:** Module 2 – Data Entry and Formatting
**Difficulty Level:** Beginner
**Estimated Time:** 60 minutes
**Prerequisites:** Lab 01 (cells, ranges, saving workbooks).

## Problem Statement

HR has a plain list of 150 employees. It is hard to read, dates and salaries are unformatted, and it cannot be printed on one page width. HR needs a polished report.

## Business Scenario

The HR head will present the employee list in a Monday meeting. Employees earning above ₹1,00,000 must stand out, headers must stay visible while scrolling, and the printout must be landscape with headers repeating on every page.

## Learning Objectives

- Enter and fix data using correct data types (text, number, date, currency).
- Apply number, date and currency formats.
- Apply fonts, alignment, borders, fills, wrap text and merge/centre.
- Convert a range to a table and apply conditional formatting.
- Freeze panes and set up print layout.

## Dataset Download

| File | Records | Link |
|---|---|---|
| employees.csv | 150 | [Download](../Allfiles/Datasets/employees.csv) |

## Dataset Column Description

| Column | Type | Description |
|---|---|---|
| EmployeeID | Text | E1001 to E1150 |
| EmployeeName | Text | Full name |
| Department | Text | Sales, HR, IT, Finance, Operations, Marketing |
| Designation | Text | Job title |
| Salary | Number | Monthly salary in INR |
| JoiningDate | Date | yyyy-mm-dd |
| Location | Text | Bengaluru, Mumbai, Delhi, Hyderabad, Chennai, Pune |

## Step 1 – Import the dataset

1. Open **employees.csv** in Excel and save as **Employee-Report.xlsx**.
1. Rename the sheet to **Data**.
1. Insert a second sheet and name it **Report**. You will build the final report on **Report** and keep **Data** as the source.

## Step 2 – Inspect the data

1. Press **Ctrl + End**. The last cell should be **G151** (150 employees + header).
1. Click a cell in column **F**. If the Formula Bar shows `2018-10-10` but the cell shows `43383`, the cell is a date with the wrong format. Fix it in Step 3.
1. Select column **E** and check the status bar: **Count** should be 150.
1. Check that Salary values are **right-aligned** (numbers). Left-aligned values would be text and must be converted.

    ![Inspecting data types in the formula bar](Linked_image_Files/lab02_step2_inspect-types.png)

## Step 3 – Apply the Excel operation

### Part A – Number, date and currency formats

1. Select **F2:F151**. Press **Ctrl + 1** (Format Cells) > **Date** > choose **14-Mar-12** style, or **Custom** and type `dd-mmm-yyyy`.
1. Select **E2:E151**. Press **Ctrl + 1** > **Currency** > Symbol: **₹ English (India)** > Decimal places: **0**.

    > *If the ₹ symbol is not listed, choose **Custom** and type `[$₹-4009] #,##,##0`.*

1. Select **A1:G1**. On **Home**, select **Bold**, **Wrap Text** and **Center** alignment, and a dark blue fill with white font.

### Part B – Format as Table

1. Click any cell in the data and press **Ctrl + T**. Confirm **My table has headers**, then **OK**.
1. On **Table Design**, name the table **tblEmployees** (left box, Table Name).
1. Choose a table style such as **Blue, Table Style Medium 2**.

    > *A table auto-expands when you add rows, keeps filter buttons, and makes formulas easier to read (you use this in Lab 03).*

### Part C – Freeze panes

1. Click cell **C2** and select **View > Freeze Panes > Freeze Panes**. The header row and the first two columns now stay visible while scrolling.

### Part D – Conditional formatting

1. Select **E2:E151**.
1. Select **Home > Conditional Formatting > Highlight Cells Rules > Greater Than**.
1. Enter **100000**, choose **Green Fill with Dark Green Text**, **OK**.
1. With the same range selected, add **Conditional Formatting > Data Bars > Gradient Fill** (blue). If you find the result too busy, remove one of the two rules with **Manage Rules**.
1. Select **D2:D151** (Designation). Choose **Conditional Formatting > Highlight Cells Rules > Text that Contains**, type **Manager**, and apply **Light Red Fill**.

    ![Conditional formatting rule for salary above 100000](Linked_image_Files/lab02_step3_conditional-format.png)

### Part E – Build the Report sheet header (merge and centre)

1. On **Report**, select **A1:G1**, then **Home > Merge & Center**. Type **Adventure Employee Information Report**. Increase font to 18, bold.
1. In **A2** type **Prepared by HR Analytics**; in **A3** type **Total employees:** and in **C3** type `=COUNTA(Data!A2:A151)`.
1. Copy **Data!A1:G151** and paste on **Report** from **A5** (Paste > Keep Source Formatting).

    > ***Tip*:** Merged cells break sorting. For titles outside the data it is fine. For tables, use **Home > Alignment settings > Horizontal: Center Across Selection** instead.

## Step 4 – Apply formulas

1. On **Data**, add a new column **H** with header **AnnualSalary**. In **H2** type:

    ```excel
    =[@Salary]*12
    ```

    The table fills the formula for all rows automatically.

1. Format **H** as currency (₹, 0 decimals).

## Step 5 – Analyse the results

1. Click the filter arrow on **Salary** > **Number Filters > Greater Than** > **100000** and read the count in the status bar (**35 of 150 records found**).
1. Clear the filter (**Data > Clear**).
1. Sort **Salary** largest to smallest. The first row should be **E1003 – Meera Sharma – ₹1,73,500**. Undo the sort or sort by **EmployeeID** again.

## Step 6 – Create visualization

1. On **Data**, select **E2:E151** and apply **Home > Conditional Formatting > Color Scales > Green - Yellow - Red Color Scale** to see the salary distribution as a heat map. (Remove the data bars first through **Manage Rules** so the two rules do not clash.)
1. For printing, go to **Page Layout**:
    - **Orientation > Landscape**.
    - **Size > A4**.
    - **Print Titles > Rows to repeat at top: $1:$1** (on Report use the header row of the pasted data).
    - **Scale to Fit > Width: 1 page**, Height: Automatic.
1. Press **Ctrl + P** and check the preview: every page shows the header row.

    ![Print preview with repeating header row](Linked_image_Files/lab02_step6_print-preview.png)

## Expected Output

| Check | Expected value |
|---|---|
| Records | 150 |
| Employees with salary > ₹1,00,000 | 35 |
| Highest salary | E1003 Meera Sharma, ₹1,73,500 |
| Average salary (status bar) | ₹73,740 |
| E1001 AnnualSalary | ₹11,16,000 (93,000 × 12) |
| Print preview | Landscape, 1 page wide, header on each page |

## Screenshot of final result

![Final formatted employee report](Linked_image_Files/lab02_final_report.png)

## Business Insights

- 35 of 150 employees (about 23%) earn above ₹1,00,000, which flags the senior and management layer.
- A formatted table with frozen headers cuts reading time in meetings and prevents misreading dates and amounts.

## Practice Questions

1. What is the shortcut that opens Format Cells?
1. Why is a date that shows as `43383` still a valid date?
1. Why can merged cells cause problems in a data table?
1. How do you make header rows repeat on every printed page?
1. What is the difference between **Freeze Top Row** and **Freeze Panes**?

<details>
<summary>Answer key</summary>

1. Ctrl + 1.
1. Excel stores dates as serial numbers; only the format is missing.
1. They block sorting, filtering and copying.
1. Page Layout > Print Titles > Rows to repeat at top.
1. Freeze Top Row locks only row 1; Freeze Panes locks rows above and columns left of the active cell.

</details>

## Lab Assignment

Create a second report sheet for the **Sales** department only (filter, copy, paste). Highlight employees who joined in or after **2022** using a conditional formatting formula rule: `=YEAR($F2)>=2022`. Print to PDF and submit the PDF and the workbook.
