---
lab:
  title: Logical functions - Employee Performance Evaluation System
  module: Module 4 - Logical Functions
  description: In this lab, you build an employee performance evaluation using IF, nested IF, AND, OR, NOT, IFS, IFERROR, COUNTIF(S), SUMIF(S) and AVERAGEIF(S).
  duration: 75 minutes
  level: 200
  islab: true
  primarytopics:
    - Excel
---

# Employee Performance Evaluation System

**Lab Number:** 04
**Lab Title:** Create an Employee Performance Evaluation System
**Module:** Module 4 – Logical Functions
**Difficulty Level:** Beginner to Intermediate
**Estimated Time:** 75 minutes
**Prerequisites:** Labs 01–03.
**Excel version note:** `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` work in Excel 2007+. **`IFS` needs Excel 2019, 2021 or Microsoft 365** (not Excel 2016 non-subscription). Use nested `IF` if `IFS` is unavailable.

## Problem Statement

HR must rate 60 sales employees every quarter based on sales target achievement, attendance and manager rating. Doing it by hand is slow and inconsistent.

## Business Scenario

Rules approved by management:

| Category | Rule (all conditions must be true) |
|---|---|
| **Outstanding** | Achievement ≥ 110% **and** Attendance ≥ 90% **and** Manager Rating ≥ 4 |
| **Good** | Achievement ≥ 90% **and** Attendance ≥ 80% **and** Manager Rating ≥ 3 |
| **Needs Improvement** | Anything else |
| **Bonus eligible** | Achievement ≥ 100% **and** Attendance ≥ 90% |

## Learning Objectives

- Write `IF`, nested `IF`, `AND`, `OR`, `NOT` and `IFS` formulas.
- Trap errors with `IFERROR`.
- Summarise with `COUNTIF(S)`, `SUMIF(S)`, `AVERAGEIF(S)`.

## Dataset Download

| File | Records | Link |
|---|---|---|
| employee_performance.csv | 60 | [Download](../Allfiles/Datasets/employee_performance.csv) |
| sales_transactions.csv (Part B) | 1,200 | [Download](../Allfiles/Datasets/sales_transactions.csv) |

## Dataset Column Description

| Column | Type | Description |
|---|---|---|
| EmployeeID | Text | E1001 – E1060 |
| EmployeeName | Text | Employee name |
| SalesTarget | Number | Quarterly target (INR) |
| SalesAchieved | Number | Actual sales (INR) |
| AttendancePct | Whole number | Attendance percentage |
| ManagerRating | Whole number | 1 (low) to 5 (high) |

## Step 1 – Import the dataset

1. Open **employee_performance.csv** and save as **Performance-Evaluation.xlsx**.
1. Convert to a table (**Ctrl + T**) named **tblPerf**, and rename the sheet **Performance**.

## Step 2 – Inspect the data

1. Check the count: **60 rows**.
1. Select **AttendancePct** and read Min and Max on the status bar (right-click the status bar to enable them): values are between 76 and 100.
1. Confirm there are no blanks in any column (**F5 > Special > Blanks**).

## Step 3 – Apply the Excel operation

1. Add column **Achievement** with `=[@SalesAchieved]/[@SalesTarget]` and format as **0.0%**.
1. Apply a conditional-formatting **Icon Set** (3 traffic lights) to **Achievement** to see who is above or below target.

## Step 4 – Apply formulas

### Part A – IF, AND, OR, NOT

| New column | Formula | Meaning |
|---|---|---|
| TargetMet | `=IF([@Achievement]>=1,"Yes","No")` | Simple IF |
| BonusEligible | `=IF(AND([@Achievement]>=1,[@AttendancePct]>=90),"Yes","No")` | Both conditions must be true |
| Flag | `=IF(OR([@AttendancePct]<80,[@ManagerRating]<=2),"Review","OK")` | Any condition true |
| NotPoor | `=IF(NOT([@ManagerRating]<=2),"Yes","No")` | Reverses TRUE/FALSE |

### Part B – Nested IF and IFS

Add column **Category** with a nested IF:

```excel
=IF(AND([@Achievement]>=1.1,[@AttendancePct]>=90,[@ManagerRating]>=4),"Outstanding",
 IF(AND([@Achievement]>=0.9,[@AttendancePct]>=80,[@ManagerRating]>=3),"Good","Needs Improvement"))
```

Add column **Category_IFS** (Excel 2019+/365):

```excel
=IFS(AND([@Achievement]>=1.1,[@AttendancePct]>=90,[@ManagerRating]>=4),"Outstanding",
     AND([@Achievement]>=0.9,[@AttendancePct]>=80,[@ManagerRating]>=3),"Good",
     TRUE,"Needs Improvement")
```

> ***Why*:** Excel tests conditions **in order** and stops at the first one that is true. The most strict rule (Outstanding) must come first. `TRUE` at the end is a catch-all.

### Part C – IFERROR

Add **AchievementSafe**: `=IFERROR([@SalesAchieved]/[@SalesTarget],0)`. If a target is ever 0 or blank, the result is 0 instead of `#DIV/0!`.

### Part D – Summary with COUNTIF/SUMIF/AVERAGEIF

On a new sheet **Summary**, type:

| Metric | Formula |
|---|---|
| Outstanding count | `=COUNTIF(tblPerf[Category],"Outstanding")` |
| Good count | `=COUNTIF(tblPerf[Category],"Good")` |
| Needs Improvement count | `=COUNTIF(tblPerf[Category],"Needs Improvement")` |
| Bonus eligible count | `=COUNTIF(tblPerf[BonusEligible],"Yes")` |
| High performers with rating ≥ 4 and attendance ≥ 90 | `=COUNTIFS(tblPerf[ManagerRating],">=4",tblPerf[AttendancePct],">=90")` |
| Total sales of Outstanding | `=SUMIF(tblPerf[Category],"Outstanding",tblPerf[SalesAchieved])` |
| Avg sales of rating 5 | `=AVERAGEIF(tblPerf[ManagerRating],5,tblPerf[SalesAchieved])` |
| Avg achieved, rating ≥ 4, attendance ≥ 90 | `=AVERAGEIFS(tblPerf[SalesAchieved],tblPerf[ManagerRating],">=4",tblPerf[AttendancePct],">=90")` |

> ***Why*:** `COUNTIF` counts cells matching **one** condition; `COUNTIFS` handles **several** (all must be true). The same idea applies to `SUMIF(S)` and `AVERAGEIF(S)`. Note that in `SUMIFS` and `AVERAGEIFS` the sum/average range comes **first**; in `SUMIF` and `AVERAGEIF` it comes **last**.

### Part E – SUMIFS on sales data

1. Open **sales_transactions.csv** (or use your Lab 03 workbook which already has NetSales). In an empty sheet, list the regions **North, South, East, West** in **A2:A5**.
1. In **B2** type `=SUMIFS(tblSales[NetSales],tblSales[Region],A2)` and fill down.
1. In **C2** type `=COUNTIFS(tblSales[Region],A2)`.
1. For the complete monthly summary left from Lab 03, list months 1–12 in **E2:E13** and use `=SUMIFS(tblSales[NetSales],tblSales[Month],E2)`.

## Step 5 – Analyse the results

1. Compare **Category** and **Category_IFS**. They must be identical. Test with `=SUMPRODUCT(--(tblPerf[Category]<>tblPerf[Category_IFS]))`, which should return **0**.
1. Filter **Category = Outstanding** and note who they are.
1. Check that the three category counts add to **60**.

## Step 6 – Create visualization

1. Select the three category names and counts on the **Summary** sheet.
1. Insert a **Clustered Column** chart titled **Employee Performance Distribution**.
1. Add data labels (**Chart Elements > Data Labels**).

    ![Performance distribution column chart](Linked_image_Files/lab04_step6_category-chart.png)

## Expected Output

| Metric | Expected value |
|---|---|
| Outstanding | 10 |
| Good | 18 |
| Needs Improvement | 32 |
| Target met (Achievement ≥ 100%) | 38 |
| Bonus eligible | 18 |
| Total achieved vs total target | ₹3,52,01,365 vs ₹3,41,00,000 |

Part E (net sales by region):

| Region | Net Sales | Orders |
|---|---|---|
| North | 1,46,06,054.50 | 305 |
| South | 1,92,86,136.00 | 314 |
| East | 1,40,88,106.50 | 274 |
| West | 1,71,15,952.50 | 307 |

## Screenshot of final result

![Final evaluation sheet with categories](Linked_image_Files/lab04_final_evaluation.png)

## Business Insights

- Only 10 of 60 employees (17%) are Outstanding, while 32 (53%) need improvement, even though 38 met their target. Many fail on attendance or manager rating, not on sales.
- The team as a whole beat its target (103%), so team-level success can hide individual weaknesses.
- South leads regional net sales (₹1.93 crore); East trails with the fewest orders (274).

## Practice Questions

1. What is the difference between `IF(AND(...))` and `IF(OR(...))`?
1. Why must the strictest rule come first in a nested IF?
1. Write a formula that gives "Pass" if score ≥ 40, else "Fail".
1. Which argument comes first in `SUMIFS`: the sum range or the criteria range?
1. What does `COUNTIF(range,">=90")` count?

<details>
<summary>Answer key</summary>

1. AND needs all conditions true; OR needs at least one.
1. Excel stops at the first true condition; a loose rule first would catch everyone.
1. `=IF(A2>=40,"Pass","Fail")`.
1. The sum range.
1. Cells with a value of 90 or more.

</details>

## Lab Assignment

Add a **Bonus** column: Outstanding gets 15% of SalesAchieved × 0.02, Good gets 8% of SalesAchieved × 0.02 and others 0. Use nested IF or IFS. Add a summary showing the total bonus per category with `SUMIF`. Explain in 3 lines how you would change the rules if attendance should count for 40% of the rating.
