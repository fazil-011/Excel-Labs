---
lab:
  title: Text and date functions - Clean customer information and calculate experience
  module: Module 6 - Text and Date Functions
  description: In this lab, you clean 80 customer records with text functions and calculate employee experience with date functions.
  duration: 75 minutes
  level: 200
  islab: true
  primarytopics:
    - Excel
---

# Clean customer information and calculate employee experience

**Lab Number:** 06
**Lab Title:** Clean Customer Information and Calculate Employee Experience
**Module:** Module 6 – Text and Date Functions
**Difficulty Level:** Intermediate
**Estimated Time:** 75 minutes
**Prerequisites:** Labs 01–05.
**Excel version note:** `CONCAT` and `TEXTJOIN` need Excel 2019+/365 (use `&` or `CONCATENATE` in Excel 2016). `DATEDIF` works in all versions but is a hidden function: it does not appear in the formula suggestions.

## Problem Statement

A customer file has names in mixed case with extra spaces, phone numbers in three formats, uppercase emails and coded product IDs. HR also needs years of service for each employee.

## Business Scenario

The marketing team wants to send a campaign, but names such as `  ananya   JOSHI ` and phones like `+91-9201828849` would look unprofessional or fail to import. HR needs experience figures for the annual appraisal.

## Learning Objectives

- Clean text with `TRIM`, `CLEAN`, `UPPER`, `LOWER`, `PROPER`, `SUBSTITUTE`.
- Extract parts of text with `LEFT`, `RIGHT`, `MID`, `LEN`, `FIND`, `SEARCH`.
- Join text with `&`, `CONCAT` and `TEXTJOIN`.
- Use `TODAY`, `NOW`, `DATE`, `DAY`, `MONTH`, `YEAR`, `DATEDIF`, `EOMONTH`, `NETWORKDAYS`, `WORKDAY`.

## Dataset Download

| File | Records | Link |
|---|---|---|
| customers_raw.csv | 80 | [Download](../Allfiles/Datasets/customers_raw.csv) |
| employees.csv | 150 | [Download](../Allfiles/Datasets/employees.csv) |

## Dataset Column Description

**customers_raw.csv**

| Column | Description |
|---|---|
| CustomerID | C0001 – C0080 |
| FullName | Messy: mixed case, extra leading/trailing/inner spaces |
| Email | Mixed upper/lower case |
| Phone | `9552087856`, `95520 87856` or `+91-9552087856` |
| City | City name |
| SignupDate | Signup date |
| ProductCode | Pattern `CAT-NNNN-CC` (category, number, country) |

**employees.csv:** see Lab 02.

## Step 1 – Import the dataset

1. Open **customers_raw.csv** and save as **Text-Date-Lab.xlsx**. Rename the sheet **Customers**.
1. Format the data as a table named **tblCust**.
1. Add a second sheet **Employees** and import **employees.csv** (copy and paste the data, make it table **tblEmp**).

## Step 2 – Inspect the data

1. Click **B2** (`SURESH KUMAR` with spaces). In a spare cell type `=LEN(B2)` and compare with the visible length. A longer number means hidden extra spaces.
1. Look at **Phone**: note the three formats.
1. Look at **ProductCode** `ELEC-0798-UK`: the pattern is 4 letters, hyphen, 4 digits, hyphen, 2 letters.

## Step 3 – Apply the Excel operation

Plan the new columns on **Customers** (add headers in H onwards):

| Header | Purpose |
|---|---|
| CleanName | Proper case, single spaces |
| FirstName | First word of the name |
| LastName | Remaining words |
| CleanEmail | Lower case |
| CleanPhone | Last 10 digits only |
| CategoryCode, ProductNo, Country | Parts of ProductCode |
| Label | Combined text for mailing |

## Step 4 – Apply formulas

### Part A – Text cleaning (customers)

| Column | Formula | What it does |
|---|---|---|
| CleanName | `=PROPER(TRIM(CLEAN([@FullName])))` | `CLEAN` removes non-printing characters; `TRIM` removes extra spaces (keeps one between words); `PROPER` capitalises each word |
| FirstName | `=LEFT([@CleanName],FIND(" ",[@CleanName])-1)` | `FIND` locates the first space; `LEFT` takes characters before it |
| LastName | `=MID([@CleanName],FIND(" ",[@CleanName])+1,100)` | `MID` starts after the space |
| CleanEmail | `=LOWER(TRIM([@Email]))` | Email in lower case |
| CleanPhone | `=RIGHT(SUBSTITUTE(SUBSTITUTE([@Phone],"+91-","")," ",""),10)` | `SUBSTITUTE` removes the country code and spaces; `RIGHT` keeps 10 digits |
| CategoryCode | `=LEFT([@ProductCode],4)` | First 4 characters |
| ProductNo | `=MID([@ProductCode],6,4)` | Start at position 6, take 4 |
| Country | `=RIGHT([@ProductCode],2)` | Last 2 characters |
| Label | `=[@CleanName]&" ("&[@City]&")"` | `&` joins text |

Extras to try:

```excel
=TEXTJOIN(" | ",TRUE,[@CleanName],[@CleanEmail],[@CleanPhone])
=CONCAT([@FirstName]," ",[@LastName])
=SEARCH("kumar",[@CleanName])
=FIND("Kumar",[@CleanName])
```

> *`FIND` is case-sensitive; `SEARCH` is not, and `SEARCH` allows wildcards. Both give `#VALUE!` if the text is not found: wrap with `IFERROR`.*

Check **C0001**: CleanName **Suresh Kumar**, FirstName **Suresh**, LastName **Kumar**, CleanEmail **suresh.kumar@outlook.com**, CleanPhone **9552087856**, CategoryCode **ELEC**, ProductNo **0798**, Country **UK**.

> **Convert formulas to values** when finished: copy the new columns, **Paste Special > Values** into a final sheet **CleanCustomers**, then delete helper columns.

### Part B – Date functions (employees)

Add columns to **tblEmp**:

| Header | Formula | Meaning |
|---|---|---|
| JoinYear | `=YEAR([@JoiningDate])` | Year |
| JoinMonth | `=MONTH([@JoiningDate])` | Month number |
| JoinDay | `=DAY([@JoiningDate])` | Day |
| YearsOfService | `=DATEDIF([@JoiningDate],DATE(2024,12,31),"y")` | Complete years until 31 Dec 2024 |
| ExperienceText | `=DATEDIF([@JoiningDate],DATE(2024,12,31),"y")&" yrs "&DATEDIF([@JoiningDate],DATE(2024,12,31),"ym")&" months"` | Years and months |
| ProbationEnd | `=EDATE([@JoiningDate],6)` | 6 months after joining |
| FirstPayday | `=EOMONTH([@JoiningDate],0)` | Last day of joining month |
| WorkingDaysFirstMonth | `=NETWORKDAYS([@JoiningDate],[@FirstPayday])` | Weekdays excluding Sat/Sun |
| ReviewDate | `=WORKDAY([@JoiningDate],90)` | 90 working days after joining |

> ***Why a fixed date?*** `TODAY()` changes every day, so your answers would differ from this lab's expected values. Use `DATE(2024,12,31)` for the lab, then replace it with `TODAY()` for a live version. `NOW()` also returns the time.

Format JoinDate-based columns (ProbationEnd, FirstPayday, ReviewDate) as `dd-mmm-yyyy`.

## Step 5 – Analyse the results

1. Check **E1001** (joined 10-Oct-2018): YearsOfService = **6**, ExperienceText = **6 yrs 2 months**, FirstPayday = **31-Oct-2018**.
1. Verify `LEN` of every **CleanPhone** is 10: in a spare cell, `=COUNTIF(tblCust[CleanPhone],"??????????")` should return **80**. (Phone is text; if it appears as a number, use `=SUMPRODUCT(--(LEN(tblCust[CleanPhone])=10))`.)
1. Find the employee with the longest service: sort **YearsOfService** descending.

## Step 6 – Create visualization

1. In a new area of the **Employees** sheet, build a small summary using `=COUNTIFS(tblEmp[YearsOfService],">=5")` and similar for 0–2, 3–4 and 5+ years.
1. Insert a **Clustered Column** chart of the three counts titled **Employees by Years of Service**.

    ![Before and after text cleaning](Linked_image_Files/lab06_step6_before-after.png)

## Expected Output

| Check | Expected |
|---|---|
| Customers cleaned | 80 rows; no leading or trailing spaces |
| `LEN` of the name for C0001 | 12 after cleaning (`Suresh Kumar`); equal or longer before |
| CleanPhone length | 10 for all 80 |
| E1001 service as of 31-Dec-2024 | 6 yrs 2 months |
| Employees processed | 150 |

## Screenshot of final result

![Final cleaned customers and experience columns](Linked_image_Files/lab06_final_clean-and-experience.png)

## Business Insights

- Clean, standardised names, emails and phones prevent bounced emails and duplicate customers in the CRM.
- Splitting `ProductCode` lets marketing segment by category and country without retyping data.
- Years-of-service data helps HR plan long-service awards and appraisal cycles.

## Practice Questions

1. What is the difference between `TRIM` and `CLEAN`?
1. What does `=MID("Excel2024",6,4)` return?
1. Why does `DATEDIF` not show in the formula autocomplete?
1. What is the difference between `NETWORKDAYS` and `WORKDAY`?
1. Why use `DATE(2024,12,31)` instead of `TODAY()` in a lab?

<details>
<summary>Answer key</summary>

1. TRIM removes extra spaces; CLEAN removes non-printing characters.
1. `2024`.
1. It is a legacy function kept for compatibility, but it works.
1. NETWORKDAYS counts working days between two dates; WORKDAY returns a date after N working days.
1. TODAY changes daily so results would not match fixed answers.

</details>

## Lab Assignment

Create a **birthday list** for the students dataset (`students.csv`): extract month and day of birth, use `=TEXT(DOB,"mmmm")` to show the month name and sort by month. Create a column **Age** using `DATEDIF` as of 31-Dec-2024. Submit the workbook.
