---
lab:
  title: Project 2 - HR Analytics
  module: Module 13 - Real-World Projects
  description: In this capstone project, you analyse headcount, department distribution, attrition, salary and performance for 300 employees and build an HR analytics dashboard.
  duration: 150 minutes
  level: 300
  islab: true
  primarytopics:
    - Excel
---

# Project 2 – HR Analytics

**Lab Number:** 14
**Lab Title:** HR Analytics – Attrition and Workforce Dashboard
**Module:** Module 13 – Real-World Projects
**Difficulty Level:** Advanced
**Estimated Time:** 150 minutes
**Prerequisites:** Labs 04, 08, 10, 12.

## Problem Statement

The company lost employees during the year and does not know who is leaving, from which department, or why. HR leaders need a workforce dashboard to guide retention decisions.

## Business Scenario

The HR Director asks: *How many employees do we have? How are they distributed? What is our attrition rate and where is it highest? How does pay vary? What drives people to leave?*

## Learning Objectives

- Calculate headcount, distribution, attrition rate and average salary.
- Create bands (age, income, tenure) with formulas.
- Compare attrition across groups with pivots.
- Build an HR dashboard and write careful, data-backed insights.

## Dataset Download

| Dataset | Link |
|---|---|
| **Practice dataset (provided, synthetic):** hr_sample.csv, 300 rows | [Download](../Allfiles/Datasets/hr_sample.csv) |
| **Real-world option:** IBM HR Analytics Employee Attrition (Kaggle) | https://www.kaggle.com/datasets/pavansubhasht/ibm-hr-analytics-attrition-dataset |

> *The IBM dataset is a fictional dataset created by IBM data scientists and is not included here: download it yourself from Kaggle. The provided file uses a subset of the same column names so that every formula in this lab also works on the Kaggle file. In the IBM file, `MonthlyIncome` is in US dollars.*

## Dataset Column Description (Data dictionary)

| Column | Type | Description |
|---|---|---|
| EmployeeNumber | Whole number | Unique employee ID (1–300) |
| Age | Whole number | Age in years (21–58) |
| Attrition | Text | **Yes** = left the company, **No** = still employed |
| Department | Text | Sales, Research & Development, Human Resources |
| JobRole | Text | Role in department |
| Gender | Text | Male / Female |
| MonthlyIncome | Whole number | Monthly salary (currency units) |
| YearsAtCompany | Whole number | Tenure in years |
| JobSatisfaction | Whole number | 1 (low) to 4 (high) |
| PerformanceRating | Whole number | 3 or 4 |
| OverTime | Text | Yes / No |
| DistanceFromHome | Whole number | Kilometres/miles from home |

## Required Excel functions

`COUNTA`, `COUNTIF(S)`, `AVERAGE`, `AVERAGEIFS`, `IF`, `IFS`, `LOOKUP`, `ROUND`, PivotTables, calculated fields, slicers, conditional formatting.

## Step 1 – Import the dataset

1. Open **hr_sample.csv**, save as **HR-Analytics.xlsx**, convert to table **tblHR** and rename the sheet **Data**.

## Step 2 – Inspect the data

1. Count rows: **300**. Check for blanks (**F5 > Special > Blanks**): none.
1. Check **Attrition** has only Yes and No: filter the column.
1. Check ranges with the status bar: Age 21–58, JobSatisfaction 1–4.

## Step 3 – Apply the Excel operation

Add helper columns:

| Column | Formula |
|---|---|
| AttritionFlag | `=IF([@Attrition]="Yes",1,0)` |
| AgeGroup | `=IFS([@Age]<30,"<30",[@Age]<40,"30-39",[@Age]<50,"40-49",TRUE,"50+")` |
| IncomeBand | `=IFS([@MonthlyIncome]<=5000,"<=5,000",[@MonthlyIncome]<=10000,"5,001-10,000",[@MonthlyIncome]<=15000,"10,001-15,000",TRUE,">15,000")` |
| TenureBand | `=IFS([@YearsAtCompany]<=2,"0-2 yrs",[@YearsAtCompany]<=5,"3-5 yrs",[@YearsAtCompany]<=10,"6-10 yrs",TRUE,"10+ yrs")` |

> *AttritionFlag lets a pivot compute the **attrition rate as the average** of the flag (1 = left).*

![Helper columns](Linked_image_Files/p2_step3_helper-columns.png)

## Step 4 – Apply formulas

| KPI | Formula | Expected |
|---|---|---|
| Headcount | `=COUNTA(tblHR[EmployeeNumber])` | 300 |
| Leavers | `=SUM(tblHR[AttritionFlag])` | 39 |
| Attrition rate | `=B3/B2` | 13.0% |
| Average monthly income | `=ROUND(AVERAGE(tblHR[MonthlyIncome]),0)` | 11,106 |
| Overtime attrition rate | `=AVERAGEIFS(tblHR[AttritionFlag],tblHR[OverTime],"Yes")` | 19.0% |
| No-overtime attrition rate | `=AVERAGEIFS(tblHR[AttritionFlag],tblHR[OverTime],"No")` | 9.7% |

## Step 5 – Analyse the results

Create pivots, always with **AttritionFlag (Average)** formatted as **0.0%** next to **Count of EmployeeNumber**:

| Pivot | Rows | Result |
|---|---|---|
| Headcount by department | Department | R&D 135, Sales 126, HR 39 |
| Attrition by department | Department | Sales 18 of 126 (14.3%); R&D 16 of 135 (11.9%); HR 5 of 39 (12.8%) |
| Attrition by overtime | OverTime | Yes 20 of 105 (19.0%); No 19 of 195 (9.7%) |
| Attrition by age group | AgeGroup | <30: 9 of 81 (11.1%); 30–39: 17 of 73 (23.3%); 40–49: 4 of 78 (5.1%); 50+: 9 of 68 (13.2%) |
| Attrition by income band | IncomeBand | ≤5,000: 4 of 31 (12.9%); 5,001–10,000: 11 of 98 (11.2%); 10,001–15,000: 11 of 95 (11.6%); >15,000: 13 of 76 (17.1%) |
| Attrition by satisfaction | JobSatisfaction | 1: 15.5%; 2: 17.6%; 3: 5.5%; 4: 13.4% |
| Average income by department | Department | Sales 11,769; R&D 10,709; HR 10,247 |
| Gender mix | Gender | Female 155, Male 145 |
| Performance | PerformanceRating | Rating 3: 203, Rating 4: 97 |

![HR pivots](Linked_image_Files/p2_step5_pivots.png)

## Step 6 – Create visualization

Build the **HR Analytics Dashboard**:

1. KPI cards: Headcount, Leavers, Attrition %, Average Income.
1. Charts: department doughnut (headcount), clustered column (attrition % by department), column (attrition % overtime vs not), column (attrition % by age group), bar (average income by department).
1. Slicers: Department, Gender, OverTime, AgeGroup.
1. Conditional formatting: highlight attrition rates above the company rate of 13.0%.

    ![HR dashboard](Linked_image_Files/p2_step6_dashboard.png)

## Expected Output

KPI row 300 / 39 / 13.0% / 11,106 and charts matching the pivot table results above.

## Screenshot of final result

![Final HR dashboard](Linked_image_Files/p2_final_dashboard.png)

## Business Insights

1. **Overtime is the clearest driver:** employees who work overtime leave at about twice the rate (19.0% vs 9.7%).
1. **Sales has the highest departmental attrition** (14.3%) while also being the highest paid on average.
1. **Employees aged 30–39 leave most** (23.3%) and 40–49 least (5.1%).
1. **Caution:** several groups have fewer than 80 employees; one or two leavers move the percentage a lot. Treat differences as hypotheses to test, not proof. The provided dataset is synthetic and built for learning.
1. Recommendation: pilot an overtime-reduction plan in Sales and run stay interviews with the 30–39 group.

## Interview questions based on this project

1. How do you calculate attrition rate in Excel?
1. Why is it better to use an AttritionFlag (1/0) than counting text?
1. How would you present a finding when the group size is small?
1. What other data would you want to explain *why* people leave?
1. How would you build an age-group column in Excel 2016 without `IFS`?

## Practice Questions

1. What is the formula for the attrition rate of the Sales department only?
1. What does `AVERAGEIFS(tblHR[AttritionFlag],tblHR[Department],"Sales")` return?
1. Why should dashboards show both counts and percentages?
1. How can a slicer for Gender change the attrition KPI?
1. What is a calculated field used for in this project?

<details>
<summary>Answer key</summary>

1. `=COUNTIFS(Department,"Sales",Attrition,"Yes")/COUNTIF(Department,"Sales")`, giving 18/126 = 14.3%.
1. 14.3%.
1. A high percentage on a tiny group is misleading without the count.
1. The pivot behind the KPI is connected to the slicer, so values update.
1. To compute a ratio of totals, such as leavers ÷ headcount, inside the pivot.

</details>

## Lab Assignment

Create an **HR one-page brief** with three charts, five insights and two recommendations. Add a **tenure-band** chart and discuss whether new joiners (0–2 years) leave more. Submit the workbook and the brief.
