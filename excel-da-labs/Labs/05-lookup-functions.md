---
lab:
  title: Lookup functions - Employee Information Retrieval System
  module: Module 5 - Lookup Functions
  description: In this lab, you build an employee lookup tool using VLOOKUP, HLOOKUP, XLOOKUP, INDEX and MATCH, approximate matching and IFERROR so that entering an employee ID retrieves the employee's details.
  duration: 75 minutes
  level: 200
  islab: true
  primarytopics:
    - Excel
---

# Employee Information Retrieval System

**Lab Number:** 05
**Lab Title:** Build an Employee Information Retrieval System
**Module:** Module 5 – Lookup Functions
**Difficulty Level:** Intermediate
**Estimated Time:** 75 minutes
**Prerequisites:** Labs 01–04.
**Excel version note:** `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH` work in **all** versions. **`XLOOKUP` needs Excel 2021 or Microsoft 365**; it is not in Excel 2016 or 2019. Use `INDEX + MATCH` if you have an older version.

## Problem Statement

HR staff scroll through 150 rows to find one employee. They need a form where typing an Employee ID instantly shows name, department, designation, salary, location and joining date.

## Business Scenario

The help desk receives dozens of calls per day: *"Which department is E1050 in? What is the salary grade?"* A lookup form saves time and avoids errors.

## Learning Objectives

- Use `VLOOKUP` (exact and approximate) and `HLOOKUP`.
- Use `INDEX` + `MATCH` and understand why it is more flexible.
- Use `XLOOKUP` (Excel 2021 / 365).
- Handle `#N/A` with `IFERROR` or the `if_not_found` argument.

## Dataset Download

| File | Records | Link |
|---|---|---|
| employees.csv | 150 | [Download](../Allfiles/Datasets/employees.csv) |

## Dataset Column Description

See Lab 02. Columns: EmployeeID, EmployeeName, Department, Designation, Salary, JoiningDate, Location.

## Step 1 – Import the dataset

1. Open **employees.csv** and save as **Employee-Lookup.xlsx**.
1. Convert the data to table **tblEmployees** (**Ctrl + T**), format **JoiningDate** as `dd-mmm-yyyy` and **Salary** as ₹.
1. Rename the sheet **Data** and add a sheet named **Lookup**.

## Step 2 – Inspect the data

1. Check that **EmployeeID** is unique: select the column, then **Home > Conditional Formatting > Highlight Cells Rules > Duplicate Values**. Nothing should highlight.
1. Note that **EmployeeID is the first (leftmost) column**. This matters: `VLOOKUP` can only return columns to the **right** of the search column.
1. Look up one record manually for later checking: **E1050 → Ananya Gupta, IT, Senior Software Engineer, ₹1,08,000, Chennai, 30-Nov-2022**.

## Step 3 – Apply the Excel operation

Design the form on the **Lookup** sheet:

| Cell | Content |
|---|---|
| A1 | Employee Information Retrieval System (merge A1:C1, bold) |
| A3 | Enter Employee ID: |
| B3 | **E1050** (input cell, yellow fill, border) |
| A5:A10 | Employee Name, Department, Designation, Salary, Location, Joining Date |
| B5:B10 | formulas (below) |

![Lookup form layout](Linked_image_Files/lab05_step3_form-layout.png)

## Step 4 – Apply formulas

### Part A – VLOOKUP (exact match)

In **B5**:

```excel
=VLOOKUP($B$3,tblEmployees,2,FALSE)
```

| Argument | Meaning |
|---|---|
| `$B$3` | Value to find (locked so it does not move when copied) |
| `tblEmployees` | The table to search; the first column is searched |
| `2` | Return the value from the 2nd column (EmployeeName) |
| `FALSE` | Exact match. **Always use FALSE for IDs.** |

Change the column number to **3** (Department), **4** (Designation), **5** (Salary), **7** (Location), **6** (JoiningDate) in B6:B10 respectively. Expected: **Ananya Gupta, IT, Senior Software Engineer, 108000, Chennai, 30-Nov-2022**.

> ***Common mistake*:** Using `TRUE` or omitting the last argument gives approximate matching, which returns **wrong results** on unsorted data.

### Part B – Make the column number dynamic with MATCH

1. Type the header names exactly as in the table into **C5:C10** (EmployeeName, Department, Designation, Salary, Location, JoiningDate) or use the table headers.
1. Replace the formula in **B5** with:

    ```excel
    =VLOOKUP($B$3,tblEmployees,MATCH(C5,tblEmployees[#Headers],0),FALSE)
    ```

    `MATCH` finds the position of the header, so inserting a column does not break the lookup.

### Part C – INDEX + MATCH

```excel
=INDEX(tblEmployees[EmployeeName],MATCH($B$3,tblEmployees[EmployeeID],0))
```

`MATCH(...,0)` returns the row position (50 for E1050); `INDEX` returns the value at that position of the chosen column.

**Look-left demonstration:** To find an ID from a name, VLOOKUP fails (the ID is left of the name). Use:

```excel
=INDEX(tblEmployees[EmployeeID],MATCH("Ananya Gupta",tblEmployees[EmployeeName],0))
```

Expected: **E1050**.

### Part D – XLOOKUP (Excel 2021 / 365 only)

```excel
=XLOOKUP($B$3,tblEmployees[EmployeeID],tblEmployees[EmployeeName],"Not found")
```

Return multiple columns at once (spills in 365):

```excel
=XLOOKUP($B$3,tblEmployees[EmployeeID],tblEmployees[[EmployeeName]:[Location]],"Not found")
```

> ***Why XLOOKUP*:** It defaults to exact match, looks in any direction, returns a custom message with `if_not_found`, and can return several columns.

### Part E – HLOOKUP

1. On a new sheet **HLookupDemo** type this small horizontal table in **A1:D2**:

    | | Sales | HR | IT |
    |---|---|---|---|
    | Budget | 500000 | 300000 | 800000 |

1. In **A5** type `=HLOOKUP("IT",A1:D2,2,FALSE)`. Result: **800000**. `HLOOKUP` searches the **first row** and returns from a row below.

### Part F – Approximate match (salary grade)

1. On the **Lookup** sheet, type a grade table in **F2:G6**:

    | Min Salary | Grade |
    |---|---|
    | 0 | D |
    | 40000 | C |
    | 70000 | B |
    | 100000 | A |

1. In **B11** (label *Salary Grade* in A11):

    ```excel
    =VLOOKUP(B8,$F$3:$G$6,2,TRUE)
    ```

    > *Approximate match (TRUE) needs the first column **sorted ascending**. It returns the largest value that is **less than or equal to** the salary.*

1. Expected for E1050 (₹1,08,000): **A**. For E1001 (₹93,000): **B**.

### Part G – Error handling with IFERROR

1. Type **E9999** in **B3**. Every lookup shows `#N/A` (the ID does not exist).
1. Wrap each formula:

    ```excel
    =IFERROR(VLOOKUP($B$3,tblEmployees,2,FALSE),"Employee not found")
    ```

1. Reset **B3** to **E1050**.

## Step 5 – Analyse the results

1. Test five IDs: E1001, E1003, E1050, E1100, E1150. Compare each with the **Data** sheet.
1. Use **Data > Data Validation > List** on **B3** with Source `=tblEmployees[EmployeeID]` if your Excel allows it (otherwise name the range `EmpIDs` and use `=EmpIDs`). This makes **B3** a dropdown (more on validation in Lab 08).
1. Confirm that the VLOOKUP, INDEX+MATCH and XLOOKUP results are identical.

## Step 6 – Create visualization

1. Add **conditional formatting** to the Salary cell **B8**: green fill if above 100000, using **Home > Conditional Formatting > Highlight Cells Rules > Greater Than**.
1. Add a **Photo-ID style card**: box with border around A5:B11, bold labels. Hide gridlines (**View > Gridlines** off).

    ![Final lookup form for E1050](Linked_image_Files/lab05_step6_final-card.png)

## Expected Output

| Field | E1050 | E1001 | E1003 |
|---|---|---|---|
| Name | Ananya Gupta | Farhan Kumar | Meera Sharma |
| Department | IT | Marketing | IT |
| Designation | Senior Software Engineer | Marketing Manager | IT Manager |
| Salary | ₹1,08,000 | ₹93,000 | ₹1,73,500 |
| Location | Chennai | Mumbai | Delhi |
| Joining date | 30-Nov-2022 | 10-Oct-2018 | 16-Nov-2020 |
| Grade | A | B | A |

## Screenshot of final result

![Final Employee Information Retrieval System](Linked_image_Files/lab05_final_lookup.png)

## Business Insights

- The help desk answers a query in seconds. One lookup form replaces repeated manual scrolling.
- Grading by salary band with approximate match makes policy changes easy: edit the table, not the formulas.

## Practice Questions

1. Why must the last argument of `VLOOKUP` be `FALSE` for employee IDs?
1. Why can `VLOOKUP` not return the EmployeeID when you search by name?
1. Which function can replace `VLOOKUP` in Excel 2016?
1. What does `MATCH(10,A1:A5,0)` return if 10 is in A3?
1. How do you show "Not found" when XLOOKUP finds nothing?

<details>
<summary>Answer key</summary>

1. Approximate matching on unsorted data returns wrong rows.
1. It only looks right of the first column.
1. `INDEX` + `MATCH`.
1. 3.
1. Use the 4th argument: `XLOOKUP(x,range,result,"Not found")`.

</details>

## Lab Assignment

Extend the form: enter an **Employee Name** (instead of an ID) and return ID, department and salary using `INDEX + MATCH`. Add a second section that shows the **number of years of service** for the ID (use `=DATEDIF(joining,TODAY(),"y")`, covered in Lab 06) and a message *"Eligible for long-service award"* if 5 or more years.
