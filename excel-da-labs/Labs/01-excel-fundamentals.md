---
lab:
  title: Excel fundamentals - Student Information Management
  module: Module 1 - Excel Fundamentals
  description: In this hands-on lab, you open a 100-record student dataset, learn the Excel interface, navigate with the Name Box and keyboard shortcuts, use relative, absolute and mixed cell references, and save the workbook in different file formats.
  duration: 45 minutes
  level: 100
  islab: true
  primarytopics:
    - Excel
---

# Student Information Management

**Lab Number:** 01
**Lab Title:** Student Information Management
**Module:** Module 1 – Excel Fundamentals
**Difficulty Level:** Beginner
**Estimated Time:** 45 minutes
**Prerequisites:** None. Any Excel version (2016, 2019, 2021 or Microsoft 365) works.

## Problem Statement

A college office keeps student marks in a CSV file. The staff must open the file, understand its structure, calculate each student's total and percentage, and save the result as a proper Excel workbook.

## Business Scenario

You are a data assistant in the college administration office. The exam cell sends you `students.csv` with 100 students. Your manager wants a clean `.xlsx` workbook where every student has **Total Marks** and **Percentage**, and wants to know how Excel stores and references data so that the sheet can be reused next semester.

## Learning Objectives

In this lab, you learn how to:

- Identify the workbook, worksheet, rows, columns and cells.
- Use the Ribbon, Formula Bar and Name Box.
- Select cells, ranges and entire columns with the mouse and keyboard.
- Write formulas that use relative, absolute and mixed references.
- Save a workbook as `.xlsx`, `.xls` and `.csv`, and explain the difference.

## Dataset Download

| File | Records | Link |
|---|---|---|
| students.csv | 100 | [Download](../Allfiles/Datasets/students.csv) |

> *Dataset type: synthetic sample generated for this course. It contains no real student data.*

## Dataset Column Description

| Column | Type | Description |
|---|---|---|
| StudentID | Text | Unique ID, S001 to S100 |
| StudentName | Text | Full name |
| Gender | Text | Male / Female |
| Course | Text | BCA, BE-CSE, BE-ECE, BSc Data Science |
| Semester | Whole number | 1 to 6 |
| City | Text | Home city |
| DateOfBirth | Date | yyyy-mm-dd |
| Phone | Text/number | 10-digit mobile number |
| Email | Text | College email |
| MathsMarks | Whole number | Out of 100 |
| ProgrammingMarks | Whole number | Out of 100 |
| DatabaseMarks | Whole number | Out of 100 |

## Step 1 – Import the dataset

1. Start Excel and select **Open > Browse**.
1. Change the file type filter (bottom right of the dialog) to **All Files (\*.\*)** or **Text Files**.
1. Select **students.csv** and select **Open**. Excel opens the CSV directly in a worksheet.

    ![Open dialog with All Files filter](Linked_image_Files/lab01_step1_open-csv.png)

1. Select **File > Save As > Browse**, set **Save as type** to **Excel Workbook (\*.xlsx)**, name it **Student-Information.xlsx** and select **Save**.

    > ***Why*:** A `.csv` file stores only text and values. It cannot store formulas, formatting or multiple sheets. Saving as `.xlsx` keeps all of them.

1. Rename the sheet tab: right-click **students** at the bottom > **Rename** > type **Students**.

## Step 2 – Inspect the data

1. Identify the parts of the Excel window and write them in a notebook:

    | Part | Where to find it | Purpose |
    |---|---|---|
    | Title bar | Top | Shows the file name |
    | Ribbon / tabs | Under the title bar | Commands grouped as Home, Insert, Formulas, Data, View |
    | Name Box | Left of the Formula Bar | Shows the active cell address; type an address to jump there |
    | Formula Bar | Top of the grid | Shows or edits the content of the active cell |
    | Column headers | A, B, C ... | Columns |
    | Row headers | 1, 2, 3 ... | Rows |
    | Sheet tabs | Bottom left | Worksheets in the workbook |
    | Status bar | Bottom | Count, Sum, Average of the selected cells |

    ![Excel window with numbered parts](Linked_image_Files/lab01_step2_interface-map.png)

1. Click cell **B2**. Notice the Name Box shows **B2** and the Formula Bar shows **Manoj Iyer**.
1. In the Name Box type **L101** and press **Enter**. You jump to the last cell of the data.
1. Press **Ctrl + Home** to go back to A1. Press **Ctrl + Shift + End** to select from A1 to the last used cell (**L101**).
1. In the Name Box type **J2:L101** and press **Enter** to select the three marks columns. Look at the status bar: it shows **Average**, **Count** and **Sum** of the selection.
1. Double-click the boundary between column headers **A** and **B** (or select all with **Ctrl + A** and use **Home > Format > AutoFit Column Width**) to make every column readable.
1. Count: the dataset has **100 records** (rows 2 to 101) and **12 columns** (A to L).

    > *A **cell** is one box (B2). A **range** is a group of cells (J2:L101). The **active cell** is the one with the thick green border.*

## Step 3 – Apply the Excel operation

In this step you format the sheet so it is easy to read.

1. Select row 1, then on **Home**, select **Bold** and a light fill colour.
1. Select **View > Freeze Panes > Freeze Top Row** so the headers stay visible while scrolling.
1. Select column **G** (DateOfBirth), then on **Home > Number**, choose **Short Date** if it shows serial numbers.
1. Select column **H** (Phone). If it shows as `9.89E+09`, widen the column or set the number format to **Number** with 0 decimals.
1. Press **Ctrl + S** to save.

## Step 4 – Apply formulas

### Part A – Relative references

1. In cell **M1** type **TotalMarks**.
1. In cell **M2** type:

    ```excel
    =SUM(J2:L2)
    ```

    > ***Why*:** `SUM` adds the three subject marks. `J2:L2` is a **range**: J2 to L2.

1. Press **Enter**. For S001 (Manoj Iyer): 48 + 46 + 89 = **183**.
1. Select **M2**, then double-click the **fill handle** (small square at the bottom-right corner of the cell). The formula copies down to **M101**.
1. Click **M3** and look at the Formula Bar: it shows `=SUM(J3:L3)`. The row number changed automatically. This is a **relative reference**.

    ![Fill handle copying formula down](Linked_image_Files/lab01_step4_fill-handle.png)

### Part B – Absolute references

1. In cell **O1** type **MaxMarks** and in **P1** type **300** (three subjects × 100).
1. In cell **N1** type **Percentage**.
1. In **N2** type:

    ```excel
    =M2/$P$1
    ```

1. Copy **N2** down to **N101**. Click **N3**: the formula is `=M3/$P$1`. The `$` signs lock the reference to **P1**. This is an **absolute reference**.
1. Select **N2:N101**, then on **Home > Number**, select **Percent Style** (%) and set **1 decimal**.
1. Press **F4** while typing a reference to cycle between `P1`, `$P$1`, `P$1` and `$P1`.

    > ***Common mistake*:** If you forget the `$`, the reference moves to P2, P3 ... and the formula shows `#DIV/0!` or wrong results after row 2.

### Part C – Mixed references

1. Add a new sheet named **RefPractice**.
1. Type the numbers **1 to 10** in **B1:K1** and **1 to 10** in **A2:A11**.
1. In **B2** type:

    ```excel
    =$A2*B$1
    ```

1. Copy **B2** right to **K2**, then select **B2:K2** and copy down to **K11**.

    > *In `$A2` the column is locked and the row moves. In `B$1` the row is locked and the column moves. This single formula builds the whole 10 × 10 multiplication table.*

## Step 5 – Analyse the results

1. Go back to the **Students** sheet and select **M2:M101**. Read **Average** on the status bar.
1. Sort check: click any cell in column **M**, then select **Data > Sort Largest to Smallest**. Note the top student. Press **Ctrl + Z** to undo the sort.
1. Confirm your answers against the **Expected Output** section.

## Step 6 – Create visualization

1. Select **StudentName** (B1:B11) and **TotalMarks** (M1:M11) using **Ctrl** to multi-select (only the first 10 students).
1. Select **Insert > Charts > Clustered Column**.
1. Add a chart title **Total Marks – First 10 Students**.

    > *A simple column chart compares one number across a few categories. You learn more chart types in Lab 09.*

## Expected Output

| Check | Expected value |
|---|---|
| Number of records | 100 (A2:L101) |
| S001 – TotalMarks | 183 |
| S001 – Percentage | 61.0% |
| Highest TotalMarks | S084, Kavya Verma, **262** |
| `$A2*B$1` table, cell K11 | 100 |

Save formats to try (**File > Save As**):

| Format | Extension | Stores formulas/formatting? | Use when |
|---|---|---|---|
| Excel Workbook | .xlsx | Yes | Default for modern Excel |
| Excel 97-2003 | .xls | Yes, limited (65,536 rows) | Old software only |
| CSV (Comma delimited) | .csv | No, values only | Exchange with other systems |

## Screenshot of final result

![Final Student-Information workbook](Linked_image_Files/lab01_final_student-workbook.png)

## Business Insights

- The college can see student totals and percentages without manual calculation.
- Using `$P$1` for max marks means that if the exam becomes out of 120 per subject, you change **one cell** and all percentages update.

## Practice Questions

1. What is the address of the cell in column F, row 25?
1. What does `Ctrl + Shift + End` do?
1. Write a formula that returns the average of one student's three marks.
1. Why does `=M2/P1` give a wrong answer when copied to row 3?
1. Name two things a `.csv` file cannot store that an `.xlsx` file can.

<details>
<summary>Answer key</summary>

1. F25.
1. Selects from the active cell to the last used cell of the data.
1. `=AVERAGE(J2:L2)`.
1. P1 is relative and becomes P2, an empty cell. Use `$P$1`.
1. Formulas, formatting, charts, multiple sheets.

</details>

## Lab Assignment

Add a column **Result** that is **Pass** if all three marks are 40 or more. (Hint: you will learn `IF` and `AND` in Lab 04; for now, use `=MIN(J2:L2)` in a helper column and check which values are below 40.) Then save a second copy of the workbook as CSV and write 3 lines explaining what was lost.
