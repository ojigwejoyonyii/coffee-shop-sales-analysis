# ☕  COFFEE SHOP SALES ANALYSIS

![cover page](https://github.com/ojigwejoyonyii/coffee-shop-sales-analysis/blob/main/Image/Cover_page.jpg)

---

## Table of Contents

- [1. 📌 Project Overview](#1--project-overview)
 
- [2. 🎯 Business Problem](#2--business-problem)
  
- [3. 🏆 Project Objectives](#3--project-objectives)
  
- [4. 🔍 Project Scope](#4-project-scope)
  
- [5. 📂 Dataset Overview](#5--dataset-overview)
  
- [6. 🧹 Data Preparation, Cleaning and Validation](#️⃣6--data-preparation-cleaning-and-validation)
  
  - [6.1 Exploratory checks](#61-exploratory-checks)
    
  - [6.2 Cleaning steps](#62-cleaning-steps)
   
  - [6.3 Calculated columns](#63-calculated-columns)
   
  - [6.4 Validation](#64-validation)
    
- [7. ⚙️ Analytical Methodology](#7--analytical-methodology)
 
- [8. 📊 Key Performance Indicators](#8--key-performance-indicators)
- 
- [9. 📈 Analysis Performed](#️⃣9--analysis-performed)
 
- [10. 📊 Dashboard Development](#10--dashboard-development)
  
- [11. 💡 Key Findings and Insights](#11--key-findings-and-insights)
  
- [12. 🚀 Recommendations](#12--recommendations)
  
- [13. ⚠️ Project Limitations](#13--project-limitations)
  
- [14. 🛠️ Tools](#14--tools)
  
- [15. 🔄 Project Workflow](#15--project-workflow)
  
- [16. 📑 Workbook Structure](#16--workbook-structure)
 
- [17. ✅ Conclusion](#17--conclusion)

---

 ## 1. 📌 Project Overview

This project analyses point-of-sale transaction data from a coffee shop business operating three store locations: **Astoria**, **Hell's Kitchen** and **Lower Manhattan**. The dataset covers **149,116 transactions** recorded over the six months from January to June 2023, generating total revenue of **$698,812.33** from **214,470 units** sold.

The raw transaction records were cleaned and validated in Microsoft Excel, enriched with calculated columns (Day, Month, Hour and Revenue), summarized with PivotTables, and presented in an interactive dashboard with slicers. The analysis examines how revenue varies by store location, product category and type, month, day of the week and hour of the day.

---

## 2. 🎯 Business Problem

The business records every sale but had no consolidated view of how its stores, products and trading periods contribute to revenue. Without this view, management cannot answer questions such as:

- Which store locations and products drive the most revenue, and which are underperforming?
- When are customers buying: which months, weekdays and hours are busiest and quietest?
- Is revenue growing, and is growth consistent across all stores?
- Where should staffing, stock and promotional effort be concentrated?

This project converts raw transaction records into structured, reliable insight that supports those decisions.

---

## 3. 🏆 Project Objectives

- Prepare a clean, validated dataset from the raw transaction records.
- Define and calculate the key performance indicators (KPIs) that describe overall sales performance.
- Compare revenue across store locations, product categories and product types.
- Identify sales patterns across months, days of the week and hours of the day.
- Build an interactive dashboard that lets stakeholders filter and explore the results.
- Translate the findings into practical, evidence-based recommendations.

---

## 4.🔍 Project Scope

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

---

## 5. 📂 Dataset Overview

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

---

## 🧹 6. Data Preparation, Cleaning and Validation

### 6.1 Exploratory checks

- **Row and column count:** The dataset contains 149,116 transaction rows (excluding the header) and 11 columns. The column count was confirmed with `COUNTA` across the header row (`=COUNTA(A1:K1)`).
- **Time coverage:** Using the filter arrows on the date field, the data was found to cover a single year, 2023, and only six months: January, February, March, April, May and June.
- **Product categories:** Filtering the product category column showed nine distinct categories were sold.
- **Data validity:** To validate the `TransactionQty` and `UnitPrice` columns, the `MIN` function was used to identify the lowest values and check for invalid negative or zero entries. The formulas `=MIN(Coffee_Shop_Sales[TransactionQty])` and `=MIN(Coffee_Shop_Sales[UnitPrice])` returned 1 and 0.8, respectively. These results confirm that both columns contain only positive minimum values, with no negative or zero entries detected.

### 6.2 Cleaning steps

| Step | Action | Result |
| --- | --- | --- |
| 1 | Renamed all column headers to PascalCase (for example, `transaction_id` became `TransactionId`) | Consistent, readable field names |
| 2 | Checked the Transaction ID column for duplicates (Data tab > Remove Duplicates). It is the only field that must be unique. | 0 duplicates found |
| 3 | Used the filter arrow on every column to check for null cells, blanks and error values | None found |
| 4 | Converted the range into an Excel Table named `Coffee_Shop_Sales`, with headers | Formulas, PivotTables and ranges update automatically as data is added |
| 5 | Formatted `UnitPrice` as currency | Correct monetary display |

### 6.3 Calculated columns

| Column | Formula | Purpose |
| --- | --- | --- |
| Day | `=TEXT(TransactionDate,"dddd")` | Day of the week as text (for example, Sunday) |
| Month | `=TEXT(TransactionDate,"mmmm")` | Month name as text (for example, January) |
| Hour | `=TEXT(TransactionTime,"h AM/PM")` | Hour of the transaction (for example, 7 AM) |
| Revenue | `=PRODUCT(UnitPrice,TransactionQty)` | Revenue per transaction (unit price × quantity) |

### 6.4 Validation

After cleaning, the following checks were run against the final data:

- Row count is unchanged between the raw and cleaned sheets (149,116).
- Transaction IDs are unique, and no field contains empty values.
- Revenue equals `UnitPrice × TransactionQty` on every row (maximum difference: $0.00).
- Every PivotTable (month, weekday, hour, category, store location and product type) reconciles to the same grand total of **$698,812.33**.

## 7. ⚙️ Analytical Methodology

The analysis is descriptive and uses a consistent approach across every dimension:

- **Aggregation:** Revenue (and, where relevant, units and transaction counts) is summed by each dimension using PivotTables built on the `Coffee_Shop_Sales` table.
- **Distinct counts:** The PivotTables were added to the Excel Data Model so that distinct counts of store locations and product categories could be calculated, which standard PivotTables cannot do.
- **Share of total:** Each segment's revenue is expressed as a percentage of total revenue to allow fair comparison.
- **Ranking:** Segments are ranked to identify top and bottom performers.
- **Trend analysis:** Month-on-month change and quarterly totals are used to describe movement over time.
- **Cross-checking:** Totals from each PivotTable are reconciled to the master total, and key results were independently recalculated from the cleaned data.
- **Visualization:** Results are charted and combined into one interactive dashboard with slicers.

## 8. 📊 Key Performance Indicators

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

---

## 9. 📈 Analysis Performed

| Analysis Area | Category / Period | Key Results | Key Insight |
|---|---|---|---|
| **Store Location** | Hell's Kitchen | Revenue: $236,511.17; Share: 33.84%; Transactions: 50,735; Units: 71,737; Avg. Transaction Value: $4.66 | Recorded the highest revenue and led monthly sales throughout the six-month period. |
| **Store Location** | Astoria | Revenue: $232,243.91; Share: 33.23%; Transactions: 50,599; Units: 70,991; Avg. Transaction Value: $4.59 | Recorded the highest sales in tea and drinking chocolate among the three stores. |
| **Store Location** | Lower Manhattan | Revenue: $230,057.25; Share: 32.92%; Transactions: 47,782; Units: 71,742; Avg. Transaction Value: $4.81 | Recorded the fewest transactions but the highest average transaction value and most units sold. |
| **Weekday Performance** | Monday | Revenue: $101,677.28 (14.55%) | Recorded the highest revenue of the week. |
| **Weekday Performance** | Saturday | Revenue: $96,894.48 (13.87%) | Recorded the lowest revenue of the week. Overall, sales were evenly distributed across the week. |
| **Hourly Performance** | 8 AM–10 AM | 36.7% of total revenue | Identified the strongest sales period, indicating a morning peak. |
| **Hourly Performance** | 7 AM–10 AM | 45.8% of total revenue | Nearly half of total revenue was generated during the morning hours. |
| **Hourly Performance** | 8 PM | Revenue: $2,935.64 (0.42%) | Recorded the lowest hourly revenue as sales tapered off in the evening. |
| **Product Category** | Coffee | Revenue: $269,952.45 (38.63%); Units: 89,250 | Contributed the highest share of total revenue. |
| **Product Category** | Tea | Revenue: $196,405.95 (28.11%); Units: 69,737 | Ranked second in revenue contribution. |
| **Product Category** | Bakery | Revenue: $82,315.64 (11.78%); Units: 23,214 | Ranked third in revenue contribution. |
| **Product Category** | Drinking Chocolate | Revenue: $72,416.00 (10.36%); Units: 17,457 | Ranked fourth in revenue contribution. |
| **Product Category** | Remaining Five Categories | Combined revenue: $77,722.29 (11.12%) | Coffee, tea, bakery and drinking chocolate collectively generated approximately 88.9% of total revenue. |
| **Product Type** | Barista Espresso | Revenue: $91,406.20 (13.08%) | Ranked first among product types by revenue. |
| **Product Type** | Brewed Chai Tea | Revenue: $77,081.95 (11.03%) | Ranked second among product types by revenue. |
| **Product Type** | Hot Chocolate | Revenue: $72,416.00 (10.36%) | Ranked third among product types by revenue. |
| **Product Type** | Gourmet Brewed Coffee | Revenue: $70,034.60 (10.02%) | Ranked fourth among product types by revenue. |
| **Product Type** | Top Four Product Types | 44.49% of total revenue | The four leading product types accounted for a substantial share of total revenue. |
| **Half-Year Performance** | Q1 (Jan–Mar 2023) | Revenue: $256,657.61 (36.7%); Units: 78,826; Transactions: 54,902 | Generated less revenue than Q2. |
| **Half-Year Performance** | Q2 (Apr–Jun 2023) | Revenue: $442,154.72 (63.3%); Units: 135,644; Transactions: 94,214 | Revenue increased by 72.3% compared with Q1. |
| **Half-Year Performance** | H1 2023 Total | Revenue: $698,812.33; Units: 214,470; Transactions: 149,116 | Average daily revenue was $3,860.84. Year-on-year comparison was not possible because the dataset covered only six months. |
| **Monthly Performance** | January | Revenue: $81,677.74; Units: 24,870; Transactions: 17,314 | Established the starting-month revenue benchmark. |
| **Monthly Performance** | February | Revenue: $76,145.19; Change: -6.8% | Revenue declined month-on-month, partly reflecting February's shorter duration. |
| **Monthly Performance** | March | Revenue: $98,834.68; Change: +29.8% | Revenue recovered strongly after February's decline. |
| **Monthly Performance** | April | Revenue: $118,941.08; Change: +20.3% | Revenue continued to increase. |
| **Monthly Performance** | May | Revenue: $156,727.76; Change: +31.8% | Recorded the highest month-on-month growth during the period. |
| **Monthly Performance** | June | Revenue: $166,485.88; Change: +6.2% | Recorded the highest monthly revenue, 103.8% above January. |

---

## 10. 📊 Dashboard Development

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

![Coffee shop sales performance dashboard](https://github.com/ojigwejoyonyii/coffee-shop-sales-analysis/blob/main/Image/coffe_shop_sales_analysis_dashboard.PNG)

---

## 11. 💡 Key Findings and Insights

- **Revenue totals $698,812.33** from 214,470 units across 149,116 transactions in six months, an average of $3,860.84 per day.
- **Strong, consistent growth:** monthly revenue doubled between January and June (+103.8%), with Q2 revenue 72.3% above Q1. June is the best month and accounts for 23.8% of all revenue.
- **Stores perform almost identically:** Hell's Kitchen leads (33.8%) but only 2.8% separates the top and bottom store, so no store is a clear underperformer on total revenue.
- **Different stores, different customers:** Lower Manhattan has the lowest traffic but the highest spend per transaction ($4.81), while Hell's Kitchen dominates coffee bean sales.
- **Coffee and tea are the business:** together they provide 66.7% of revenue, and four product types provide 44.5%.
- **A pronounced morning peak:** 7–10 AM generates 45.8% of revenue. Revenue falls by half at 11 AM and stays flat through the afternoon.
- **Very weak late-evening trade:** the 8 PM hour earns only 0.42% of revenue, and 6 AM earns just 3.1%.
- **The trading week is balanced:** weekday revenue varies by less than 5%, with Monday strongest and Saturday weakest.
- **A long tail of small categories:** Branded, Loose Tea, Flavours and Packaged Chocolate together contribute 5.4% of revenue.

---

## 12. 🚀 Recommendations

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

---

## 13. ⚠️ Project Limitations

- **Short time window:** Only six months (January–June 2023) are included. Seasonality, year-on-year growth and the cause of the January–June increase cannot be established.
- **Revenue, not profit:** No cost, margin, discount or refund data is available, so product and store rankings reflect revenue only.
- **No customer or staff data:** There is no customer ID or employee field, so repeat-purchase behaviour, customer segments and salesperson performance cannot be analysed.
- **Transaction definition:** Each transaction ID contains a single product line, so "average transaction value" is the average value per sale line and not per customer basket.
- **Currency assumed:** Amounts are formatted as currency in the workbook, and are presented here in US dollars ($).
- **Descriptive analysis only:** The work identifies patterns but does not test statistical significance or forecast future sales. Differences between stores in particular are small.
- **Manual workflow:** Cleaning and calculated columns are performed in Excel and must be refreshed manually when new data is added.

---

## 14. 🛠️ Tools

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

---

## 15. 🔄 Project Workflow

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

---

## 16. 📑 Workbook Structure

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

---

## 17. ✅ Conclusion

The project turned 149,116 raw transaction records into a clean, validated dataset, a set of reconciled KPIs and an interactive dashboard. The analysis shows a healthy and fast-growing business: revenue doubled between January and June 2023, the three stores perform almost equally, and coffee and tea are the dominant products. Sales peak sharply in the morning and are weak late in the evening.

The clearest opportunities are to match staffing and stock to the morning peak, build demand in the flat afternoon period, test whether the earliest and latest trading hours justify their cost, and copy store-specific strengths across locations. The main limitation is the six-month window, which means that growth and seasonality cannot yet be separated. Extending the data to a full year, and adding cost and customer information, would make the next phase of analysis considerably more powerful.

