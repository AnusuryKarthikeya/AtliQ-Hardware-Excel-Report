# AtliQ Hardware — Excel Report Generation

Excel-based reporting workflow that uses **Power Query** to import and transform AtliQ Hardware's raw CSV sales data into four business reports: **Customer Net Sales Performance**, **Market Performance**, **P&L Year**, and **P&L Month**.

## Overview

AtliQ Hardware's sales data is distributed as raw CSV exports, with a **fiscal year running September–August**. This workbook uses Power Query to import, clean, and join the CSVs, then presents the results as formatted PivotTable-based reports with slicers matching the standard filter set: **region, market, customer, division**.

## Data Sources (CSV files)

Loaded as two query groups in Power Query — **dimension** and **fact**:

| Group | Table | Contents | Rows Loaded |
|---|---|---|---|
| dimension | `dim_customer` | Customer master data | 189 |
| dimension | `dim_market` | Market/region master data | 23 |
| dimension | `dim_product` | Product master data | 298 |
| dimension | `dim_date` | Calendar/fiscal date table | 1,066 |
| fact | `fact_sales_monthly` | Monthly sales transactions (quantity, pricing, deductions, cost measures) | 7,99,962 |
| fact | `ns_targets_2021` | 2021 net sales targets by market, used for the Market Performance vs Target report | 276 |

## Tech Stack

- **Microsoft Excel**
- **Power Query** (Get & Transform Data) — CSV import, cleaning, joins
- **Power Pivot / Data Model** — relationships between `fact_sales_monthly`, `ns_targets_2021`, and the dimension tables
- **PivotTables & Slicers** — report layout and filtering (region, market, customer, division)
- **Excel formulas / DAX measures** — Net Sales Amount, COGS, Gross Margin, GM %

## Project Structure

```
atliq-hardware-excel-reports/
├── README.md
├── data/
│   └── raw/
│       ├── dim_customer.csv
│       ├── dim_market.csv
│       ├── dim_product.csv
│       ├── dim_date.csv
│       ├── fact_sales_monthly.csv
│       └── ns_targets_2021.csv
├── AtliQ_Hardware_Reports.xlsx      # main workbook — Power Query + Data Model + PivotTables
└── exports/                          # PDF exports of each report page
    ├── Customer_Net_Sales.pdf
    ├── Market_Performance.pdf
    ├── P&L_Year.pdf
    └── P&L_Month.pdf
```

## Setup

### 1. Place the source CSVs

Save all raw CSV files into `data/raw/` (or point Power Query at wherever they're stored).

### 2. Open the workbook and load the queries

Open `AtliQ_Hardware_Reports.xlsx`. Each source CSV is connected through its own Power Query query (**Data tab → Queries & Connections**), organized into a `dimension` group (`dim_customer`, `dim_market`, `dim_product`, `dim_date`) and a `fact` group (`fact_sales_monthly`, `ns_targets_2021`). Each query:
- Sets correct data types (dates, numbers, text)
- Derives a **Fiscal Year** column (Sep–Aug) from `dim_date`
- Cleans/renames columns as needed

### 3. Data Model relationships

`fact_sales_monthly` and `ns_targets_2021` are linked to `dim_customer`, `dim_market`, `dim_product`, and `dim_date` via the Data Model (Power Pivot), enabling the calculated fields below to work across all four reports.

### 4. Calculated fields

- **Net Sales Amount**, **Total COGS**, and **Gross Margin** are derived measures built from the quantity, pricing, deduction, and cost columns in `fact_sales_monthly`
- **GM %** = Gross Margin ÷ Net Sales Amount
- **2021-Target** and **2021-Target %** (Market Performance report) come from `ns_targets_2021`, compared against actual Net Sales Amount by market

### 5. Refresh data

When new/updated CSVs are dropped into `data/raw/`:
- **Data tab → Refresh All**, or
- Right-click an individual query → **Refresh**

## Report Details

### Customer Net Sales Performance
Net sales by customer across fiscal years, with year-over-year growth.

- **Rows:** one per customer (e.g. Amazon, AtliQ e Store, AtliQ Exclusive, Flipkart, Croma), plus a Grand Total row
- **Columns:** 2019, 2020, 2021, 2021 vs 2020 (%)
- **Slicers:** region, market, division

### Market Performance vs Target
Net sales by market against the 2021 sales target (from `ns_targets_2021`).

- **Rows:** one per market (e.g. India, USA, China, Germany, Australia), plus a Grand Total row
- **Columns:** 2019, 2020, 2021, 2021-Target (variance in USD), 2021-Target % (variance)
- **Slicers:** region, division

### P&L Year
Annual profit & loss summary.

- **Rows (metrics):** Net Sales Amount, Total COGS, Gross Margin, GM %
- **Columns:** fiscal years (2019, 2020, 2021)
- **Slicers:** region, market, customer, division

### P&L Month
Monthly profit & loss breakdown, generated per fiscal year.

- **Rows (metrics):** Net Sales Amount, Total COGS, Gross Margin, GM %
- **Columns:** months Sep–Aug grouped under Q1–Q4, with a Grand Total column
- **Additional table:** Net Sales Comparison — YoY % by month (e.g. "21 vs 20", "20 vs 19")
- **Slicers:** region, market, division, fiscal year (report is viewed one fiscal year at a time)

## Troubleshooting

| Issue | Fix |
|---|---|
| Query fails to load | Confirm the CSV file paths in Power Query's source step still match `data/raw/` |
| `fact_sales_monthly` shows a load warning | Open the query and check the applied steps for type-conversion or column-mismatch errors before relying on the data |
| Blank/zero values after refresh | Check that new CSVs use the same column headers and date format as the originals |
| PivotTable not updating | Refresh the Data Model first (Data tab → Refresh All), then refresh the PivotTable |
| Fiscal year grouping looks wrong | Verify `dim_date`'s fiscal year column uses a Sep–Aug year boundary, not calendar year |

## License

Internal/portfolio use.
