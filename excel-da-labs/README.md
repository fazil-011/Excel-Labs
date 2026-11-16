# Microsoft Excel for Data Analytics – Hands-on Labs

Lab-by-lab, step-by-step instructions in the same format as the Microsoft PL-300 (Power BI) labs: front-matter header, numbered steps, bold UI names, notes, expected output and screenshots. Every lab ships with a dataset you can practise on.

## How to use

1. Download the repository (or this folder) and keep the folder structure unchanged.
1. Open a lab in `Labs/` and follow it from top to bottom.
1. Each lab links to its dataset in `Allfiles/Datasets/` (relative links, so they work on GitHub and locally).
1. Compare your result with the **Expected Output** section, then answer the practice questions (answers are folded under *Answer key*).

## Lab index

| Lab | Module | Title | Dataset | Time | Excel version |
|---|---|---|---|---|---|
| [01](Labs/01-excel-fundamentals.md) | 1 Fundamentals | Student Information Management | students.csv (100) | 45 min | Any |
| [02](Labs/02-data-entry-and-formatting.md) | 2 Data Entry & Formatting | Employee Information Report | employees.csv (150) | 60 min | Any |
| [03](Labs/03-basic-formulas-and-functions.md) | 3 Basic Formulas | Monthly Sales Summary | sales_transactions.csv (1,200) | 60 min | Any |
| [04](Labs/04-logical-functions.md) | 4 Logical Functions | Employee Performance Evaluation | employee_performance.csv (60) | 75 min | `IFS`: 2019/365 |
| [05](Labs/05-lookup-functions.md) | 5 Lookup Functions | Employee Information Retrieval | employees.csv (150) | 75 min | `XLOOKUP`: 2021/365 |
| [06](Labs/06-text-and-date-functions.md) | 6 Text & Date | Clean Customers, Calculate Experience | customers_raw.csv (80), employees.csv | 75 min | `CONCAT/TEXTJOIN`: 2019/365 |
| [07](Labs/07-data-cleaning.md) | 7 Data Cleaning | Clean a Messy Sales Dataset | customer_sales_messy.csv (160) | 90 min | Any |
| [08](Labs/08-sorting-filtering-and-validation.md) | 8 Sort, Filter, Validate | Inventory Management Worksheet | inventory.csv (36) | 75 min | Any |
| [09](Labs/09-charts-and-visualization.md) | 9 Charts | Sales Performance Visualization Report | sales_transactions.csv | 90 min | 2010+ |
| [10](Labs/10-pivot-tables-and-pivot-charts.md) | 10 Pivot Tables | Interactive Sales Analysis Report | sales_transactions.csv | 90 min | 2013+ (slicers/timelines) |
| [11](Labs/11-power-query.md) | 11 Power Query | Consolidate 12 Monthly Files | monthly_sales/ (12 files), salespeople.csv, sales_targets_wide.csv | 90 min | 2016+ Windows (2010/2013 add-in) |
| [12](Labs/12-advanced-excel-dashboards.md) | 12 Dashboards | Sales Analytics Dashboard | sales_transactions.csv | 120 min | 2013+ |
| [13](Labs/13-project-1-retail-sales-analytics.md) | 13 Project 1 | Retail Sales Analytics | sales_transactions.csv / Kaggle Superstore | 150 min | 2019/365 for `IFS` |
| [14](Labs/14-project-2-hr-analytics.md) | 13 Project 2 | HR Analytics | hr_sample.csv / Kaggle IBM HR | 150 min | 2019/365 for `IFS` |
| [15](Labs/15-project-3-financial-analytics.md) | 13 Project 3 | Financial Analytics | finance_monthly.csv / Kaggle Financial Sample | 150 min | Waterfall: 2016+ |
| [16](Labs/16-project-4-ecommerce-analytics.md) | 13 Project 4 | E-commerce Analytics | ecommerce_orders.csv / Kaggle E-commerce Data | 150 min | Any (2016+) |

## Excel version cheat sheet

| Feature | Excel 2016 | Excel 2019 | Excel 2021 | Microsoft 365 |
|---|---|---|---|---|
| VLOOKUP, INDEX/MATCH, SUMIFS, PivotTables | Yes | Yes | Yes | Yes |
| IFS, CONCAT, TEXTJOIN | No (2016 subscription: yes) | Yes | Yes | Yes |
| XLOOKUP, UNIQUE, dynamic arrays | No | No | Yes | Yes |
| Power Query built in | Yes (Windows) | Yes | Yes | Yes |
| Waterfall chart | Yes | Yes | Yes | Yes |
| Slicers / Timelines | Yes | Yes | Yes | Yes |

## Folder structure

```text
excel-da-labs/
├── README.md
├── SCREENSHOTS.md              # list of screenshots to capture
├── Labs/
│   ├── 01-excel-fundamentals.md … 16-project-4-ecommerce-analytics.md
│   └── Linked_image_Files/     # put your screenshots here
└── Allfiles/
    └── Datasets/               # CSV files (see Datasets/README.md)
```

## Important notes

- **Screenshots are not included.** Each lab references images such as `Linked_image_Files/lab03_step6_monthly-line.png`. Capture them while you run the lab; `SCREENSHOTS.md` lists every file name and what it should show.
- **Datasets** named in the labs are synthetic practice files generated for this course so that results are repeatable. Kaggle datasets are linked separately and must be downloaded by you; they are not bundled. See `Allfiles/Datasets/README.md`.
- Currency values use Indian rupees (₹). Indian digit grouping (1,50,000) appears in some tables; Excel shows your system's grouping.

- `generate_datasets.py` recreates every CSV in `Allfiles/Datasets/` with the same random seeds (run it from this folder with Python 3).
