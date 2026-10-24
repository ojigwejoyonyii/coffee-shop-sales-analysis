 # ☕ Coffee Shop Sales Analysis

> **Data period:** 1 January – 30 June 2023 | **Tool:** Microsoft Excel | **Prepared by:** Ojigwe Joy Onyinye

![Dashboard](images/dashboard.png)  

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Business Problem](#2-business-problem)
3. [Project Objectives](#3-project-objectives)
4. [Project Scope](#4-project-scope)
5. [Dataset Overview](#5-dataset-overview)
6. [Data Source](#6-data-source)
7. [Data Preparation, Cleaning and Validation](#7-data-preparation-cleaning-and-validation)
8. [Analytical Methodology](#8-analytical-methodology)
9. [Key Performance Indicators](#9-key-performance-indicators)
10. [Analysis Performed](#10-analysis-performed)
11. [Dashboard Development](#11-dashboard-development)
12. [Key Findings and Insights](#12-key-findings-and-insights)
13. [Recommendations](#13-recommendations)
14. [Project Limitations](#14-project-limitations)
15. [Tools](#15-tools)
16. [Project Workflow](#16-project-workflow)
17. [Workbook Structure](#17-workbook-structure)
18. [Conclusion](#18-conclusion)
19. [Appendix A: Data Dictionary](#appendix-a-data-dictionary)

---

## 1. Project Overview

This project analyses point-of-sale transaction data from a coffee shop business operating three store locations: **Astoria**, **Hell's Kitchen** and **Lower Manhattan**. The dataset covers **149,116 transactions** recorded over the six months from January to June 2023, generating total revenue of **$698,812.33** from **214,470 units** sold.

The raw transaction records were cleaned and validated in Microsoft Excel, enriched with calculated columns (Day, Month, Hour and Revenue), summarised with PivotTables, and presented in an interactive dashboard with slicers. The analysis examines how revenue varies by store location, product category and type, month, day of the week and hour of the day.

## 2. Business Problem

The business records every sale but had no consolidated view of how its stores, products and trading periods contribute to revenue. Without this view, management cannot answer questions such as:

- Which store locations and products drive the most revenue, and which are underperforming?
- When are customers buying: which months, weekdays and hours are busiest and quietest?
- Is revenue growing, and is growth consistent across all stores?
- Where should staffing, stock and promotional effort be concentrated?

This project converts raw transaction records into structured, reliable insight that supports those decisions.

## 3. Project Objectives

- Prepare a clean, validated dataset from the raw transaction records.
- Define and calculate the key performance indicators (KPIs) that describe overall sales performance.
- Compare revenue across store locations, product categories and product types.
- Identify sales patterns across months, days of the week and hours of the day.
- Build an interactive dashboard that lets stakeholders filter and explore the results.
- Translate the findings into practical, evidence-based recommendations.

## 4. Project Scope

**In scope**

- All 149,116 transactions in the dataset, from 1 January to 30 June 2023.
- All three store locations, nine product categories and 29 product types.
- Revenue-focused descriptive analysis by location, product, month, weekday and hour.
- Data cleaning, validation, KPI calculation, PivotTable analysis and dashboard development in Excel.

**Out of scope**

- Profitability analysis (the dataset contains no cost or margin data).
- Customer-level analysis such as segmentation or retention (there is no customer identifier).
- Individual salesperson or staff performance (there is no employee field).
- Forecasting, year-on-year comparison, or any period outside January–June 2023.

## 5. Dataset Overview

| Attribute | Detail |
| --- | --- |
| Records | 149,116 transactions (one product line per transaction ID) |
| Raw columns | 11 |
| Columns after cleaning | 15 (11 original, with 4 calculated columns added) |
| Period covered | 1 January 2023 – 30 June 2023 (181 trading days) |
| Store locations | 3 – Astoria, Hell's Kitchen, Lower Manhattan |
| Product categories | 9 |
| Product types | 29 |
| Distinct products (product ID) | 80 |
| Unit price range | $0.80 – $45.00 |
| Quantity per transaction | 1 – 8 units (98% of transactions contain 1 or 2 units) |

The nine product categories are Bakery, Branded, Coffee, Coffee beans, Drinking Chocolate, Flavours, Loose Tea, Packaged Chocolate and Tea.

## 6. Data Source

The data is a transactional sales extract sourced from the link provided by the organization. Each row is one sale line, identified by a unique transaction ID, with its date, time, quantity, store, product and unit price. The raw sheet was left untouched; all cleaning and calculation was carried out on a separate copy (`Cleaned_Sales_Data`) so the original data can always be traced and re-checked.

## 7. Data Preparation, Cleaning and Validation

### 7.1 Exploratory checks

- **Row and column count:** The dataset contains 149,116 transaction rows (excluding the header) and 11 columns. The column count was confirmed with `COUNTA` across the header row (`=COUNTA(A1:K1)`).
- **Time coverage:** Using the filter arrows on the date field, the data was found to cover a single year, 2023, and only six months: January, February, March, April, May and June.
- **Product categories:** Filtering the product category column showed nine distinct categories were sold.
- **Data validity:** To validate the `TransactionQty` and `UnitPrice` columns, the `MIN` function was used to identify the lowest values and check for invalid negative or zero entries. The formulas `=MIN(Coffee_Shop_Sales[TransactionQty])` and `=MIN(Coffee_Shop_Sales[UnitPrice])` returned 1 and 0.8, respectively. These results confirm that both columns contain only positive minimum values, with no negative or zero entries detected.

### 7.2 Cleaning steps

| Step | Action | Result |
| --- | --- | --- |
| 1 | Renamed all column headers to PascalCase (for example, `transaction_id` became `TransactionId`) | Consistent, readable field names |
| 2 | Checked the Transaction ID column for duplicates (Data tab > Remove Duplicates). It is the only field that must be unique. | 0 duplicates found |
| 3 | Used the filter arrow on every column to check for null cells, blanks and error values | None found |
| 4 | Converted the range into an Excel Table named `Coffee_Shop_Sales`, with headers | Formulas, PivotTables and ranges update automatically as data is added |
| 5 | Formatted `UnitPrice` as currency | Correct monetary display |

### 7.3 Calculated columns

| Column | Formula | Purpose |
| --- | --- | --- |
| Day | `=TEXT(TransactionDate,"dddd")` | Day of the week as text (for example, Sunday) |
| Month | `=TEXT(TransactionDate,"mmmm")` | Month name as text (for example, January) |
| Hour | `=TEXT(TransactionTime,"h AM/PM")` | Hour of the transaction (for example, 7 AM) |
| Revenue | `=PRODUCT(UnitPrice,TransactionQty)` | Revenue per transaction (unit price × quantity) |

### 7.4 Validation

After cleaning, the following checks were run against the final data:

- Row count is unchanged between the raw and cleaned sheets (149,116).
- Transaction IDs are unique, and no field contains empty values.
- Revenue equals `UnitPrice × TransactionQty` on every row (maximum difference: $0.00).
- Every PivotTable (month, weekday, hour, category, store location and product type) reconciles to the same grand total of **$698,812.33**.

## 8. Analytical Methodology

The analysis is descriptive and uses a consistent approach across every dimension:

- **Aggregation:** Revenue (and, where relevant, units and transaction counts) is summed by each dimension using PivotTables built on the `Coffee_Shop_Sales` table.
- **Distinct counts:** The PivotTables were added to the Excel Data Model so that distinct counts of store locations and product categories could be calculated, which standard PivotTables cannot do.
- **Share of total:** Each segment's revenue is expressed as a percentage of total revenue to allow fair comparison.
- **Ranking:** Segments are ranked to identify top and bottom performers.
- **Trend analysis:** Month-on-month change and quarterly totals are used to describe movement over time.
- **Cross-checking:** Totals from each PivotTable are reconciled to the master total, and key results were independently recalculated from the cleaned data.
- **Visualization:** Results are charted and combined into one interactive dashboard with slicers.

## 9. Key Performance Indicators

| KPI | Value | Definition | On dashboard |
| --- | --- | --- | --- |
| Total Revenue | $698,812.33 | Sum of Revenue | Yes |
| Total Units Sold | 214,470 | Sum of TransactionQty | Yes |
| Store Locations | 3 | Distinct count of StoreLocation | Yes |
| Product Categories | 9 | Distinct count of ProductCategory | Yes (labelled "Product Type") |
| Total Transactions | 149,116 | Count of TransactionId | No |
| Average Transaction Value | $4.69 | Total revenue ÷ transactions | No |
| Average Units per Transaction | 1.44 | Total units ÷ transactions | No |
| Average Daily Revenue | $3,860.84 | Total revenue ÷ 181 trading days | No |
| Product Types / Products | 29 / 80 | Distinct product types and product IDs | No |

## 10. Analysis Performed

### 10.1 Store Locations Sales Analysis

Revenue is spread almost evenly across the three stores. Hell's Kitchen leads, followed by Astoria and Lower Manhattan, but the gap between first and last is only $6,453.92 (2.8%).

| Store location | Revenue | Share | Transactions | Units | Avg. transaction value |
| --- | --- | --- | --- | --- | --- |
| Hell's Kitchen | $236,511.17 | 33.84% | 50,735 | 71,737 | $4.66 |
| Astoria | $232,243.91 | 33.23% | 50,599 | 70,991 | $4.59 |
| Lower Manhattan | $230,057.25 | 32.92% | 47,782 | 71,742 | $4.81 |

- Hell's Kitchen recorded the highest revenue in each of the six months.
- Lower Manhattan has the fewest transactions but sells the most units and has the highest average transaction value, so its customers buy more per visit.
- Product mix differs by store. Hell's Kitchen sells far more coffee beans ($18,635 versus about $10,200–$11,200 elsewhere) but far fewer branded items ($1,942 versus $5,457–$6,208). Astoria leads in tea and drinking chocolate, and Lower Manhattan leads in bakery and flavours.

![Analysis of coffee shop store locations performance](images/store-location-analysis.png)
*Image 1: Analysis of coffee shop store locations performance.*

### 10.2 Weekday and Hourly Performance Analysis

The dataset has no salesperson or employee field, so individual performance cannot be measured. This section instead covers the closest available view of operational performance: when revenue is earned by day of the week and hour of the day.

**Revenue by day of the week**

| Day | Revenue | Share |
| --- | --- | --- |
| Monday | $101,677.28 | 14.55% |
| Friday | $101,373.00 | 14.51% |
| Thursday | $100,767.78 | 14.42% |
| Wednesday | $100,313.54 | 14.36% |
| Tuesday | $99,455.94 | 14.23% |
| Sunday | $98,330.31 | 14.07% |
| Saturday | $96,894.48 | 13.87% |

Sales are very evenly distributed across the week. Monday is the strongest day and Saturday the weakest, a difference of only about 4.9%.

![Analysis of weekday performance](images/weekday-analysis.png)
*Image 2: Analysis of weekday performance.*

**Revenue by hour of the day**

| Hour | Revenue | Share | Hour | Revenue | Share |
| --- | --- | --- | --- | --- | --- |
| 6 AM | $21,900.27 | 3.13% | 2 PM | $41,304.74 | 5.91% |
| 7 AM | $63,526.47 | 9.09% | 3 PM | $41,733.10 | 5.97% |
| 8 AM | $82,699.87 | 11.83% | 4 PM | $41,122.75 | 5.88% |
| 9 AM | $85,169.53 | 12.19% | 5 PM | $40,134.31 | 5.74% |
| 10 AM | $88,673.39 | 12.69% | 6 PM | $34,286.20 | 4.91% |
| 11 AM | $46,319.14 | 6.63% | 7 PM | $28,446.68 | 4.07% |
| 12 PM | $40,192.79 | 5.75% | 8 PM | $2,935.64 | 0.42% |
| 1 PM | $40,367.45 | 5.78% | | | |

The business has a strong morning peak. The three hours from 8 AM to 10 AM bring in 36.7% of all revenue, and 7 AM to 10 AM brings in 45.8%. Revenue then halves at 11 AM and stays flat at about $40,000–$42,000 per hour through to 5 PM, before tapering off in the evening. The 8 PM hour contributes just 0.42% of revenue.

![Analysis of performance of hours of the day](images/hourly-analysis.png)
*Image 3: Analysis of performance of hours of the day.*

### 10.3 Product-Type Analysis

**By product category**

| Category | Revenue | Share | Units | Transactions |
| --- | --- | --- | --- | --- |
| Coffee | $269,952.45 | 38.63% | 89,250 | 58,416 |
| Tea | $196,405.95 | 28.11% | 69,737 | 45,449 |
| Bakery | $82,315.64 | 11.78% | 23,214 | 22,796 |
| Drinking Chocolate | $72,416.00 | 10.36% | 17,457 | 11,468 |
| Coffee beans | $40,085.25 | 5.74% | 1,828 | 1,753 |
| Branded | $13,607.00 | 1.95% | 776 | 747 |
| Loose Tea | $11,213.60 | 1.60% | 1,210 | 1,210 |
| Flavours | $8,408.80 | 1.20% | 10,511 | 6,790 |
| Packaged Chocolate | $4,407.64 | 0.63% | 487 | 487 |

Coffee and tea together generate 66.7% of revenue. Adding bakery and drinking chocolate brings the four core categories to 88.9%. The remaining five categories contribute 11.1% combined.

![Analysis of product category performance](images/product-category-analysis.png)
*Image 4: Analysis of product category performance.*

**Top product types**

| Rank | Product type | Revenue | Share |
| --- | --- | --- | --- |
| 1 | Barista Espresso | $91,406.20 | 13.08% |
| 2 | Brewed Chai tea | $77,081.95 | 11.03% |
| 3 | Hot chocolate | $72,416.00 | 10.36% |
| 4 | Gourmet brewed coffee | $70,034.60 | 10.02% |
| 5 | Brewed Black tea | $47,932.00 | 6.86% |
| 6 | Brewed herbal tea | $47,539.50 | 6.80% |
| 7 | Premium brewed coffee | $38,781.15 | 5.55% |
| 8 | Organic brewed coffee | $37,746.50 | 5.40% |
| 9 | Scone | $36,866.12 | 5.28% |
| 10 | Drip coffee | $31,984.00 | 4.58% |

The top four product types alone account for 44.5% of revenue. The best-selling individual item (by product detail) is *Sustainably Grown Organic Lg* at $21,151.75, followed closely by *Dark chocolate Lg* ($21,006.00) and *Latte Rg* ($19,112.25). The lowest-earning product types are Green beans ($1,340.00), Green tea ($1,470.75) and Organic Chocolate ($1,679.60).

![Analysis of product types performance](images/product-type-analysis.png)
*Image 5: Analysis of product types performance.*

### 10.4 Yearly Sales Analysis

The dataset covers only 2023, and only its first six months, so a true year-on-year comparison is not possible. The analysis therefore summarises the half-year and compares its two quarters.

| Period | Revenue | Share of period total | Units | Transactions |
| --- | --- | --- | --- | --- |
| Q1 (Jan–Mar 2023) | $256,657.61 | 36.7% | 78,826 | 54,902 |
| Q2 (Apr–Jun 2023) | $442,154.72 | 63.3% | 135,644 | 94,214 |
| **H1 2023 total** | **$698,812.33** | **100%** | **214,470** | **149,116** |

Second-quarter revenue is 72.3% higher than the first quarter. Average daily revenue for the half-year is $3,860.84. Because the data stops at June, the figures should not be annualised: doing so would assume the Q2 level continues all year, which the data cannot confirm.

### 10.5 Monthly Sales Analysis

| Month | Revenue | Month-on-month change | Units | Transactions |
| --- | --- | --- | --- | --- |
| January | $81,677.74 | – | 24,870 | 17,314 |
| February | $76,145.19 | -6.8% | 23,550 | 16,359 |
| March | $98,834.68 | +29.8% | 30,406 | 21,229 |
| April | $118,941.08 | +20.3% | 36,469 | 25,335 |
| May | $156,727.76 | +31.8% | 48,233 | 33,527 |
| June | $166,485.88 | +6.2% | 50,942 | 35,352 |

Revenue rose in every month except February, and June revenue is 103.8% higher than January's. February's apparent dip is largely explained by its length: it has 28 days against January's 31. On a per-day basis February ($2,719) was slightly above January ($2,635). The growth is broad-based: all three stores grew in each month from March to June, and Hell's Kitchen led in every month.

![Analysis of monthly performance](images/monthly-analysis.png)
*Image 6: Analysis of monthly performance.*

## 11. Dashboard Development

The Dashboard worksheet brings the analysis together in a single interactive view.

**Components**

- **KPI cards:** Total Revenue, Total Units Sold, Product Categories and Store Locations.
- **Charts (six):** monthly revenue (line), weekday revenue (column), hourly revenue (line), revenue by product category (column), revenue by store location (pie) and revenue by product type (bar).
- **Slicers (six):** Day, Month, Hour, Store Location, Product Category and Product Type. All slicers are connected to every PivotTable, so a single selection updates all charts together.
- **Branding:** a header banner and logo image, with a consistent colour scheme.

**Build approach**

1. Build one PivotTable per analysis on its own worksheet.
2. Create a PivotChart from each PivotTable.
3. Copy the charts to the Dashboard sheet and link them to the shared slicers via Report Connections.
4. Add the KPI cards, title and layout formatting.

**Usage example:** selecting "Hell's Kitchen" in the Store Location slicer and "Coffee" in the Product Category slicer shows that store's coffee performance by month, weekday and hour.

![Coffee shop sales performance dashboard](images/dashboard.png)
*Image 7: Coffee shop sales performance dashboard.*

## 12. Key Findings and Insights

- **Revenue totals $698,812.33** from 214,470 units across 149,116 transactions in six months, an average of $3,860.84 per day.
- **Strong, consistent growth:** monthly revenue doubled between January and June (+103.8%), with Q2 revenue 72.3% above Q1. June is the best month and accounts for 23.8% of all revenue.
- **Stores perform almost identically:** Hell's Kitchen leads (33.8%) but only 2.8% separates the top and bottom store, so no store is a clear underperformer on total revenue.
- **Different stores, different customers:** Lower Manhattan has the lowest traffic but the highest spend per transaction ($4.81), while Hell's Kitchen dominates coffee bean sales.
- **Coffee and tea are the business:** together they provide 66.7% of revenue, and four product types provide 44.5%.
- **A pronounced morning peak:** 7–10 AM generates 45.8% of revenue. Revenue falls by half at 11 AM and stays flat through the afternoon.
- **Very weak late-evening trade:** the 8 PM hour earns only 0.42% of revenue, and 6 AM earns just 3.1%.
- **The trading week is balanced:** weekday revenue varies by less than 5%, with Monday strongest and Saturday weakest.
- **A long tail of small categories:** Branded, Loose Tea, Flavours and Packaged Chocolate together contribute 5.4% of revenue.

## 13. Recommendations

> The recommendations below are based on six months of data. They are hypotheses to test, and should be confirmed against cost and staffing data before they are acted on.

1. **Align staffing and stock to the morning peak.** Concentrate staff and prepared stock from 7 to 10 AM, which produces nearly half of daily revenue, and avoid over-staffing the flat midday and afternoon period.
2. **Create demand in the afternoon.** Test an afternoon offer (for example, a bakery and drink bundle) to lift the flat 11 AM–5 PM period, which has the most room to grow.
3. **Review trading hours.** Compare the labour and operating cost of the 8 PM hour (0.42% of revenue) and the 6 AM hour (3.1%) against the revenue they earn, and consider adjusting opening or closing times.
4. **Replicate Hell's Kitchen's coffee bean performance.** Investigate why it sells roughly 1.7 to 1.8 times as many coffee beans as the other stores, and test the same display, staff recommendation or promotion elsewhere.
5. **Improve merchandise visibility in Hell's Kitchen.** Branded items sell at roughly a third of the level of the other stores, which suggests a display or placement opportunity.
6. **Grow traffic at Lower Manhattan.** Its customers already spend the most per visit, so footfall is its main lever.
7. **Protect the core range.** Keep Barista Espresso, Brewed Chai tea, Hot chocolate and Gourmet brewed coffee consistently available and high quality.
8. **Review the low-revenue range.** Packaged Chocolate, Loose Tea and Branded items earn little, so review whether their shelf space and stock investment are justified. Flavours sells often but at a low value, and may work best as an add-on to drinks.
9. **Plan capacity for continued growth,** but treat the trend with caution until a full year of data confirms how much of it is seasonal.
10. **Extend the dataset** to a full year, and add cost, discount and staffing data to move from revenue analysis to profitability analysis.

## 14. Project Limitations

- **Short time window:** Only six months (January–June 2023) are included. Seasonality, year-on-year growth and the cause of the January–June increase cannot be established.
- **Revenue, not profit:** No cost, margin, discount or refund data is available, so product and store rankings reflect revenue only.
- **No customer or staff data:** There is no customer ID or employee field, so repeat-purchase behaviour, customer segments and salesperson performance cannot be analysed.
- **Transaction definition:** Each transaction ID contains a single product line, so "average transaction value" is the average value per sale line and not per customer basket.
- **Currency assumed:** Amounts are formatted as currency in the workbook, and are presented here in US dollars ($).
- **Descriptive analysis only:** The work identifies patterns but does not test statistical significance or forecast future sales. Differences between stores in particular are small.
- **Manual workflow:** Cleaning and calculated columns are performed in Excel and must be refreshed manually when new data is added.

## 15. Tools

| Tool or feature | Use in the project |
| --- | --- |
| Microsoft Excel (.xlsb workbook) | Primary platform for cleaning, analysis and reporting |
| Excel Tables | Named table (`Coffee_Shop_Sales`) with auto-expanding references |
| Filters and Remove Duplicates | Data exploration, and checking for blanks, errors and duplicate IDs |
| Excel functions: `TEXT`, `PRODUCT`, `COUNTA` | Calculated columns and validation checks |
| PivotTables and PivotCharts | Aggregation and visual summaries |
| Excel Data Model | Distinct counts of store locations and product categories |
| Slicers | Interactive filtering across the dashboard |
| Microsoft Word | Project documentation |

## 16. Project Workflow

1. **Understand the data:** Open the raw file, then count rows and columns and identify the period and categories covered.
2. **Standardize:** Rename the headers to PascalCase and convert the range to the `Coffee_Shop_Sales` table.
3. **Validate:** Check for duplicate transaction IDs, blank cells and errors.
4. **Enrich:** Add the Day, Month, Hour and Revenue columns, and apply currency formatting.
5. **Model:** Load the data into the Data Model to enable distinct counts.
6. **Analyze:** Build PivotTables for month, weekday, hour, product category, product type and store location.
7. **Visualize:** Create a PivotChart for each analysis.
8. **Assemble:** Place the KPI cards, charts and slicers on the Dashboard sheet and connect the slicers to all PivotTables.
9. **Interpret:** Draw out the findings and recommendations.
10. **Document:** Record the method, results and limitations (this document).

## 17. Workbook Structure

The workbook (`coffe_shop_sales_analysis.xlsb`) contains 11 worksheets:

| # | Worksheet | Contents |
| --- | --- | --- |
| 1 | `Raw_Sales_Data` | Original, unmodified transaction data (11 columns) |
| 2 | `Cleaned_Sales_Data` | Cleaned table with calculated columns (15 columns) |
| 3 | `Store_Location_Distinct_Count` | PivotTable: distinct count of store locations (3) |
| 4 | `Product_Category_Distinct_Count` | PivotTable: distinct count of product categories (9) |
| 5 | Revenue by month | PivotTable and line chart of monthly revenue |
| 6 | Revenue by weekday | PivotTable and column chart of weekday revenue |
| 7 | Revenue by hour | PivotTable and line chart of hourly revenue |
| 8 | Revenue by product category | PivotTable and column chart of category revenue |
| 9 | Revenue by store location | PivotTable and pie chart of store revenue |
| 10 | Revenue by product type | PivotTable and bar chart of product type revenue |
| 11 | Dashboard | KPI cards, six charts and six slicers |

## 18. Conclusion

The project turned 149,116 raw transaction records into a clean, validated dataset, a set of reconciled KPIs and an interactive dashboard. The analysis shows a healthy and fast-growing business: revenue doubled between January and June 2023, the three stores perform almost equally, and coffee and tea are the dominant products. Sales peak sharply in the morning and are weak late in the evening.

The clearest opportunities are to match staffing and stock to the morning peak, build demand in the flat afternoon period, test whether the earliest and latest trading hours justify their cost, and copy store-specific strengths across locations. The main limitation is the six-month window, which means that growth and seasonality cannot yet be separated. Extending the data to a full year, and adding cost and customer information, would make the next phase of analysis considerably more powerful.

## Appendix A: Data Dictionary

**`Cleaned_Sales_Data` (15 columns)**

| Column | Data type | Description | Origin |
| --- | --- | --- | --- |
| `TransactionId` | Integer | Unique identifier for each transaction (sale line) | Original (`transaction_id`) |
| `TransactionDate` | Date | Date of the sale, 1 Jan – 30 Jun 2023 | Original (`transaction_date`) |
| `Day` | Text | Day of the week, for example Sunday | Calculated |
| `Month` | Text | Month name, for example January | Calculated |
| `TransactionTime` | Time | Time the sale was recorded | Original (`transaction_time`) |
| `Hour` | Text | Hour of the sale, for example 7 AM | Calculated |
| `TransactionQty` | Integer | Number of units sold in the transaction (1–8) | Original (`transaction_qty`) |
| `StoreId` | Integer | Numeric identifier of the store | Original (`store_id`) |
| `StoreLocation` | Text | Store name: Astoria, Hell's Kitchen or Lower Manhattan | Original (`store_location`) |
| `ProductId` | Integer | Numeric identifier of the product (80 distinct) | Original (`product_id`) |
| `UnitPrice` | Currency | Price per unit ($0.80 – $45.00) | Original (`unit_price`) |
| `Revenue` | Currency | `UnitPrice × TransactionQty` | Calculated |
| `ProductCategory` | Text | Top-level product group (9 categories) | Original (`product_category`) |
| `ProductType` | Text | Sub-group within a category (29 types) | Original (`product_type`) |
| `ProductDetail` | Text | Product name including size, for example Latte Rg | Original (`product_detail`) |

**Size abbreviations in `ProductDetail`:** Sm = small, Rg = regular, Lg = large.
