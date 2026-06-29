````markdown
# Retail Sales Optimization Using Excel

An end-to-end retail analytics project developed using **Microsoft Excel, Power Query, Power Pivot, DAX, PivotTables, PivotCharts, and interactive slicers**.

The project transforms raw retail transactions into a validated dimensional model and an interactive executive dashboard that supports sales, profitability, customer, product, geographic, and returns analysis.

---

## Project Overview

This project analyzes retail sales performance and identifies opportunities to improve revenue, profitability, order value, product performance, and operational efficiency.

The solution covers the complete analytics workflow:

- Data ingestion and transformation
- Data cleaning and validation
- Dimensional data modeling
- DAX measure development
- Interactive dashboard creation
- Business insight generation
- Data-quality documentation

---

## Business Problem

Retail management needs a reliable way to answer questions such as:

- How much revenue and profit has the business generated?
- Which product categories generate the highest sales and profitability?
- Which sub-categories are underperforming or generating losses?
- How do sales and profits change over time?
- What is the return rate?
- Which regions and customer segments perform best?
- Are there data-quality problems that could affect management decisions?

---

## Dataset

The project uses the **Sample Superstore** retail dataset.

The dataset contains information about:

- Orders
- Customers
- Products
- Product categories and sub-categories
- Geographic locations
- Shipping modes
- Sales
- Quantity
- Discounts
- Profit
- Returned orders

The final fact table contains **10,194 sales records**.

---

## Tools and Technologies

- Microsoft Excel
- Power Query
- Power Pivot
- DAX
- Excel Data Model
- PivotTables
- PivotCharts
- Slicers
- Conditional Formatting

---

## ETL Workflow

The data preparation process was completed using Power Query.

### Main ETL operations

1. Imported the Orders, Returns, and People tables.
2. Created connection-only raw queries.
3. Standardized column data types.
4. Applied text trimming and cleaning.
5. Validated missing and erroneous values.
6. Merged returned-order information with the sales table.
7. Added regional manager information.
8. Created dimension tables.
9. Generated surrogate keys.
10. Loaded the analytical tables into the Excel Data Model.

---

## Star Schema

The final data model follows a star-schema design.

### Fact Table

**FactSales**

Contains:

- Row ID
- Order ID
- Order Date
- Ship Date
- Customer ID
- ProductKey
- GeographyKey
- ShipModeKey
- Sales
- Quantity
- Discount
- Profit
- Returned

### Dimension Tables

- **DimProduct**
- **DimCustomer**
- **DimDate**
- **DimGeography**
- **DimShipMode**

### Relationships

```text
DimProduct[ProductKey]      1 ─── * FactSales[ProductKey]
DimCustomer[Customer ID]    1 ─── * FactSales[Customer ID]
DimGeography[GeographyKey]  1 ─── * FactSales[GeographyKey]
DimShipMode[ShipModeKey]    1 ─── * FactSales[ShipModeKey]
DimDate[Date]               1 ─── * FactSales[Order Date]
````

---

## Data Quality Checks

The project includes a dedicated data-quality validation process.

| Validation Check            |              Result | Status      |
| --------------------------- | ------------------: | ----------- |
| Raw sales rows              |              10,194 | Passed      |
| Final FactSales rows        |              10,194 | Passed      |
| Duplicate Row IDs           |                   0 | Passed      |
| Missing ProductKey values   |                   0 | Passed      |
| Missing GeographyKey values |                   0 | Passed      |
| Missing ShipModeKey values  |                   0 | Passed      |
| Unique Customer IDs         |                 804 | Passed      |
| Product ID conflicts        | 32 IDs / 64 records | Warning     |
| Returned orders             |                 296 | Information |

### Product ID Remediation

The original Product ID was not a reliable unique business key.

A total of **32 Product IDs** were associated with **64 conflicting product records**. To preserve valid product records and maintain referential integrity, a numeric surrogate key named `ProductKey` was created.

---

## DAX Measures

The following DAX measures were created:

```DAX
Total Sales :=
SUM(FactSales[Sales])
```

```DAX
Total Profit :=
SUM(FactSales[Profit])
```

```DAX
Total Orders :=
DISTINCTCOUNT(FactSales[Order ID])
```

```DAX
Total Quantity :=
SUM(FactSales[Quantity])
```

```DAX
Profit Margin % :=
DIVIDE([Total Profit], [Total Sales], 0)
```

```DAX
Average Order Value :=
DIVIDE([Total Sales], [Total Orders], 0)
```

```DAX
Returned Orders :=
CALCULATE(
    DISTINCTCOUNT(FactSales[Order ID]),
    FactSales[Returned] = "Yes"
)
```

```DAX
Return Rate % :=
DIVIDE([Returned Orders], [Total Orders], 0)
```

Additional time-intelligence measures were developed for previous-year sales, previous-year profit, and year-over-year growth.

---

## Dashboard Features

The interactive executive dashboard includes:

* Total Sales
* Total Profit
* Total Orders
* Total Quantity
* Profit Margin
* Average Order Value
* Returned Orders
* Return Rate
* Monthly Sales and Profit Trend
* Sales and Profit Margin by Category
* Profit by Sub-Category
* Year slicer
* Region slicer
* Customer Segment slicer
* Product Category slicer

---

## Key Results

| KPI                 |        Result |
| ------------------- | ------------: |
| Total Sales         | $2,326,534.35 |
| Total Profit        |   $292,296.81 |
| Total Orders        |         5,111 |
| Total Quantity      |        38,654 |
| Profit Margin       |        12.56% |
| Average Order Value |       $455.20 |
| Returned Orders     |           296 |
| Return Rate         |         5.79% |

---

## Key Business Insights

* Technology generated the highest sales and strongest profitability among the main categories.
* Furniture generated substantial sales but weaker profitability than Technology and Office Supplies.
* The overall business achieved a positive profit margin of 12.56%.
* The average order value was $455.20, indicating opportunities for cross-selling and product bundling.
* The return rate was 5.79%, requiring further analysis by product, region, category, and shipping mode.
* Product identifier conflicts demonstrated the importance of surrogate keys and referential-integrity validation.
* Monthly sales and profit patterns can support inventory planning and seasonal forecasting.

---

## Recommendations

* Review discount levels on low-margin and loss-making products.
* Investigate Furniture sub-categories with weak or negative profitability.
* Prioritize high-performing Technology products in marketing campaigns.
* Analyze returned orders to identify preventable product or shipping problems.
* Use product bundles and complementary recommendations to increase average order value.
* Continue validating Product IDs during future data refreshes.
* Use regional and customer-segment analysis to improve inventory allocation.

---

## Repository Structure

```text
retail-sales-optimization-excel/
│
├── Retail_Sales_Optimization_Final.xlsx
├── README.md
│
├── screenshots/
│   ├── executive_dashboard.png
│   ├── data_quality.png
│   └── business_insights.png
│
└── docs/
    └── project_summary.pdf
```

---

## How to Use the Excel File

1. Download `Retail_Sales_Optimization_Final.xlsx`.
2. Open the file using Microsoft Excel Desktop.
3. Enable Power Pivot if it is not already enabled.
4. Open the `Executive_Dashboard` worksheet.
5. Use the Year, Region, Segment, and Category slicers.
6. Select **Data → Refresh All** to refresh the queries and analytical model.

> The complete interactive experience requires a desktop version of Excel that supports Power Query, the Data Model, and Power Pivot.

---

## Screenshots

### Executive Dashboard

![Executive Dashboard](screenshots/executive_dashboard.png)

### Data Quality Report

![Data Quality Report](screenshots/data_quality.png)

### Business Insights

![Business Insights](screenshots/business_insights.png)

---

## Project Skills Demonstrated

* ETL development
* Data cleaning
* Data-quality validation
* Dimensional modeling
* Star-schema design
* Referential-integrity management
* Surrogate-key creation
* DAX measure development
* Time-intelligence analysis
* Dashboard design
* Business analysis
* Data storytelling

---

## Author

Developed as an Excel-based retail sales analytics and optimization portfolio project.

```
```
