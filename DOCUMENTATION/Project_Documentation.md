# ☕ Coffee Shop Sales Performance Analysis

**Maven Roasters** | Sales Data: January – June 2023


![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)
![PivotTables](https://img.shields.io/badge/Technique-Pivot%20Tables-blue)
![Dashboard](https://img.shields.io/badge/Output-Interactive%20Dashboard-orange)


**Prepared by:** Joy Onyinye

**Project Completed:** July 2026

---

## 📑 Table of Contents

- [Executive Summary](#-executive-summary)
- [Business Problem](#-business-problem)
- [Project Objectives](#-project-objectives)
- [Data Source and Dataset Description](#-data-source-and-dataset-description)
- [Data Cleaning](#-data-cleaning)
- [Data Preparation and Feature Engineering](#-data-preparation-and-feature-engineering)
- [Data Analysis — Pivot Table Summarization](#-data-analysis--pivot-table-summarization)
- [Visualization Layer](#-visualization-layer)
- [Interactive Dashboard Assembly](#-interactive-dashboard-assembly)
- [Findings](#-findings)
- [Recommendations](#-recommendations)
- [Limitations](#-limitations)
- [Conclusion](#-conclusion)

---

## 📌 Executive Summary

This report presents a sales performance analysis of **Maven Roasters**, a coffee shop chain with three locations in New York City, using transactional data from **January to June 2023**. The analysis transformed raw sales data into actionable business insights through data preparation, transformation, and interactive dashboard development in Microsoft Excel.

**Key highlights:**

| Metric | Value |
|---|---|
| 💰 Total revenue (6 months) | **$698,812.33** |
| 🏆 Top product category | **Coffee** — $269,952.45 (38.6%) |
| 📈 Highest revenue month | **June** |
| ⏰ Busiest trading hour | **10:00 a.m.** |
| 📅 Highest weekday revenue | **Monday** — $113,845.83 (16.3%) |
| 📉 Lowest weekday revenue | **Saturday** |

These findings provide valuable insight into customer purchasing behavior and offer a basis for informed decisions on inventory management, staffing, marketing, and operational planning.

---

## 🧩 Business Problem

Maven Roasters operates three coffee shop locations in New York City and generates a substantial volume of daily sales transactions. Despite the availability of transactional sales data, management lacks a structured approach for converting this data into meaningful business insights that can support strategic and operational decision-making.

Without a comprehensive analysis of sales performance, it is difficult to identify top-performing products, peak trading periods, seasonal trends, and customer purchasing patterns. This information gap limits management's ability to:

- Optimize inventory levels
- Allocate staff effectively during peak hours
- Develop targeted marketing campaigns
- Implement strategies that improve sales performance and profitability

There is therefore a need to transform the available sales data into an interactive and insightful reporting solution that enables management to monitor KPIs, evaluate performance across dimensions, and make data-driven decisions.

---

## 🎯 Project Objectives

The primary objective of this project is to analyze the sales performance of **Maven Roasters** using transactional data from **January to June 2023** and transform it into meaningful business insights.

Specific objectives:

1. Prepare and transform the raw sales dataset into a clean, structured, analysis-ready format by creating calculated fields (Revenue, Day,  Day Name, Month, Hour Label).
2. Evaluate overall sales performance — total revenue and transaction volume — over the six-month reporting period.
3. Identify monthly sales trends to determine periods of peak and low business performance.
4. Examine sales performance across product categories to identify top revenue contributors.
5. Analyze customer purchasing patterns by day of week and hour of day to identify peak and low-activity periods.
6. Develop an interactive dashboard that communicates KPIs, trends, and insights clearly.
7. Generate actionable, evidence-based recommendations for inventory, staffing, marketing, and operations.
8. Demonstrate Excel data analysis techniques — data transformation, Pivot Tables, Pivot Charts, slicers, and dashboard design — applied to a real-world business problem.

---

## 🗂️ Data Source and Dataset Description

Data source: [Google Drive link](https://drive.google.com/uc?export=download&id=1z6_Srogh8t7MCSxij_syKj8sKcntQ-Ki)

The primary data resides in the **"Transactions"** sheet — a raw, single-table extract of point-of-sale records.

| Attribute | Detail |
|---|---|
| **Record count** | 149,116 individual transaction line items |
| **Fields (11)** | `transaction_id`, `transaction_date`, `transaction_time`, `transaction_qty`, `store_id`, `store_location`, `product_id`, `unit_price`, `product_category`, `product_type`, `product_detail` |
| **Date range** | 1 January 2023 – 30 June 2023 (first half-year only) |
| **Store locations** | 3 branches — Lower Manhattan, Hell's Kitchen, Astoria |
| **Product categories** | 9 categories (e.g., Coffee, Tea, Bakery, Drinking Chocolate, Coffee Beans, Branded, Loose Tea, Flavours, Packaged Chocolate) |
| **Product types** | 29 distinct product types nested within the 9 categories |
| **Unit price range** | $0.80 – $45.00 |
| **Quantity per line** | 1 to 8 units |

---

## 🧹 Data Cleaning

The following cleaning steps were performed on the coffee shop sales dataset:

1. **Duplication of raw data** — To preserve the raw data, the sheet was duplicated (right-click → Move or Copy → Create a Copy) and renamed `transaction_clean`.
2. **Converting range to table** — The dataset was converted to a formal Excel Table (Home tab → Format as Table → "My table has headers") to enable further cleaning steps.
3. **Checking for missing values** — Used `=COUNTBLANK(range)` on each column to confirm no missing values existed.
4. **Checking for duplicate records** — Used Data → Remove Duplicates (all columns checked); returned zero duplicates.
5. **Validating data types** — Verified `transaction_date` and `transaction_time` were stored as true date/time values (not text) via cell alignment and the `HOUR()` function test.
6. **Checking numeric fields for invalid values** — Used `=MIN(range)` on `transaction_qty` 
and `unit_price`; returned 1 and 0.80 respectively, confirming no invalid (≤0) entries.
7. **Checking text fields for consistency** — Used filter dropdowns on `product_category`, `store_location`, and `product_type` to check for inconsistent entries.
8. **Formatting column headers** — Converted multi-word headers from `snake_case` to `PascalCase` (e.g., `TransactionId`, `TransactionDate`).
9. **Formatting price values** — Applied USD currency formatting to `unit_price` and `Revenue`, since the business operates in New York.

### Data Quality Check Summary

| Check performed | Result |
|---|---|
| Missing/null values across all 11 fields | None found — every field fully populated for all 149,116 rows |
| Duplicate rows (exact row duplicates) | None found |
| Duplicate `transaction_id` values | None found — `transaction_id` is a clean unique key |
| Non-positive or invalid quantity/price values | None found — all quantities and prices are positive numbers |
| Date range continuity | Continuous daily coverage from 1 Jan to 30 Jun 2023, no gaps identified |
| Row count consistency (raw vs. cleaned) | Identical (149,116 rows in both `Transactions` and `transaction_clean`) — no rows dropped or filtered |

> Because no nulls, duplicates, or invalid values were present, the cleaning stage did not need to remove or impute any records. The "cleaning" undertaken was primarily a **feature-engineering** step, described next.

---

## 🛠️ Data Preparation and Feature Engineering

A second table, `transaction_clean`, was built alongside the raw data. It carries all 11 original fields (renamed to PascalCase) plus **five derived columns**, each computed with a live
formula so it recalculates automatically if the source data changes:

| Derived column | Formula | Purpose |
|---|---|---|
| **Day** | `=DAY(cell)` | Extracts the day number from the full date |
| **DayName** | `=TEXT(cell, "dddd")` | Extracts the weekday name (text) for readability |
| **Month** | `=TEXT(cell, "mmmm")` | Extracts the month name (text) |
| **Revenue** | `=PRODUCT(qty_cell, price_cell)` | Multiplies quantity sold × unit price |
| **HourLabel** | `=TEXT(cell, "h AM/PM")` | Extracts the hour of the transaction for hourly analysis |

---

## 📊 Data Analysis — Pivot Table Summarization

Using `transaction_clean` as the source range, **seven pivot tables** were built (Insert tab → PivotTable), each condensing the 149,116-row detail table into a compact summary along one 
dimension. All summaries use **Revenue** (sum of revenue, or count of revenue as a stand-in for transaction count) as the measure.

| Summary sheet | Dimension (Row Labels) | Measure |
|---|---|---|
| Revenue per weekday | Weekday (DayName) | Sum of Revenue |
| Number of sales on weekdays | Weekday (DayName) | Count of Revenue (transaction count) |
| Revenue on hourly basis | Hour (HourLabel) | Sum of Revenue |
| Number of sales on hourly basis | Hour (HourLabel) | Count of Revenue |
| Revenue on monthly basis | Month | Sum of Revenue |
| Number of sales on monthly basis | Month | Count of Revenue |
| Revenue by each category | Product Category | Sum of Revenue |

**Build steps per pivot table:**

- **Revenue per weekday** — `DayName` → Rows, `Revenue` → Values (Sum)
- **Number of sales per weekday** — `DayName` → Rows, `Revenue` → Values (Count)
- **Revenue on hourly basis** — `HourLabel` → Rows, `Revenue` → Values (Sum)
- **Number of sales on hourly basis** — `HourLabel` → Rows, `Revenue` → Values (Count)
- **Revenue on monthly basis** — `Month` → Rows, `Revenue` → Values (Sum)
- **Number of sales on monthly basis** — `Month` → Rows, `Revenue` → Values (Count)
- **Revenue by category** — `ProductCategory` → Rows, `Revenue` → Values (Sum)

Pivot charts were then created for each table (Insert → PivotChart → Column Chart), with field buttons hidden and descriptive chart titles applied.

---

## 📈 Visualization Layer

For each pivot summary, a corresponding bar chart was created — **fourteen charts in total**: one set embedded next to its source
pivot table on each summary sheet, and a duplicate, formatted set reused on the `DASHBOARD` sheet. All charts use a simple bar/column format for easy single-dimension comparisons.


**Revenue by weekday**
 <img width="1050" height="600" alt="revenue-by-weekday" src="https://github.com/user-attachments/assets/acf88634-abf5-4042-a64c-89932d1b714c" />

*Figure 1 — Revenue by weekday, reconstructed from the "revenue per weekday" pivot table*



**Revenue by hour of day**
 <img width="1350" height="600" alt="revenue-by-hour" src="https://github.com/user-attachments/assets/efff45b3-95ef-4b95-b30e-98e719649446" />

*Figure 2 — Revenue by hour of day, reconstructed from the "Revenue on hourly basis" pivot table*



**Revenue by month**
 <img width="1050" height="600" alt="revenue-by-month" src="https://github.com/user-attachments/assets/2c268630-a0ef-4ec6-b6bd-d17eeb816941" />

*Figure 3 — Revenue by month, reconstructed from the "revenue on monthly basis" pivot table*



**Revenue by product category**
<img width="1200" height="675" alt="revenue-by-category" src="https://github.com/user-attachments/assets/b71a8037-513e-475a-b7cc-58a4c1d44f79" />

*Figure 4 — Revenue by product category, reconstructed from the "Revenue by each category" pivot table*

