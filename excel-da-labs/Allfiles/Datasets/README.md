# Datasets

## Provided practice datasets (synthetic)

Generated for this course with fixed random seeds. They contain no real people or companies, and results in the labs are computed from these exact files.

| File | Rows | Used in | Description |
|---|---|---|---|
| students.csv | 100 | Lab 01 | Student details and marks in three subjects |
| employees.csv | 150 | Labs 02, 05, 06 | Employee ID, name, department, designation, salary, joining date, location |
| sales_transactions.csv | 1,200 | Labs 03, 04 (Part E), 09, 10, 12, Project 1 | 2024 order lines: region, salesperson, category, product, quantity, price, cost, discount |
| monthly_sales/sales_2024_01.csv … _12.csv | 1,200 in total | Lab 11 | The same transactions split by month |
| salespeople.csv | 8 | Lab 11 | Salesperson → manager, team, join year |
| sales_targets_wide.csv | 8 | Lab 11 | Monthly targets in wide format (for Unpivot) |
| employee_performance.csv | 60 | Lab 04 | Target, achieved sales, attendance, manager rating |
| customers_raw.csv | 80 | Lab 06 | Messy names, emails, phones and product codes |
| customer_sales_messy.csv | 160 | Lab 07 | Duplicates, blanks, spelling variants, invalid dates, text amounts |
| customer_sales_clean_reference.csv | 150 | Lab 07 | Answer key (ground truth before the mess was added) |
| inventory.csv | 36 | Lab 08 | Products, stock, reorder level, price |
| hr_sample.csv | 300 | Project 2 | Attrition and workforce fields, using IBM-style column names |
| finance_monthly.csv | 96 | Project 3 | Budget and actual by account for 12 months |
| ecommerce_orders.csv | 1,510 | Project 4 | Order lines with customer and country (500 orders, 145 customers) |

## Kaggle datasets (download yourself)

These links come from the course plan. They have **not** been downloaded or verified in this package; Kaggle requires a free account, and dataset content or availability can change. Check each dataset's licence on its Kaggle page.

| # | Dataset | Link | Business use | Map to |
|---|---|---|---|---|
| 1 | Superstore Sales | https://www.kaggle.com/datasets/vivek468/superstore-dataset-final | Retail sales, profit and discount analysis | Labs 03, 09, 10, 12; Project 1 |
| 2 | IBM HR Analytics Employee Attrition | https://www.kaggle.com/datasets/pavansubhasht/ibm-hr-analytics-attrition-dataset | Attrition, salary and satisfaction analysis | Labs 04, 10; Project 2 |
| 3 | E-commerce Data | https://www.kaggle.com/datasets/carrie1/ecommerce-data | Orders, customers and products for an online retailer | Labs 07, 11; Project 4 |
| 4 | Financial Sample | https://www.kaggle.com/datasets/atharvaarya25/financials | Sales and profit by segment, country and product | Lab 09; extension of Project 3 |
| 5 | Walmart Sales | https://www.kaggle.com/datasets/mikhail1681/walmart-sales | Weekly store sales for trend analysis | Lab 10 optional assignment; Project 1 extension |

If you cannot access a Kaggle dataset, use the provided practice dataset in the "Map to" labs; the formulas and steps are the same.
