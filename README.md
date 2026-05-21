# 🚴 Adventure Works Cycles — Internet Sales Dashboard

> An end-to-end Excel sales analytics dashboard analysing **60,398 transactions** across **5 years**, **10 territories** and **606 products** to uncover revenue trends, top customers, and regional performance.

---

## 📌 Table of Contents

1. [Project Title](#1-project-title)
2. [Brief Summary](#2-brief-summary)
3. [Overview](#3-overview)
4. [Problem Statement](#4-problem-statement)
5. [Dataset](#5-dataset)
6. [Tools and Technologies](#6-tools-and-technologies)
7. [Methods](#7-methods)
8. [Key Insights](#8-key-insights)
9. [Dashboard / Output](#9-dashboard--output)
10. [How to Run This Project](#10-how-to-run-this-project)
11. [Results and Conclusion](#11-results-and-conclusion)
12. [Future Work](#12-future-work)

---

## 1. Project Title

### 🚴 Adventure Works Cycles — Internet Sales Performance Dashboard (2010–2014)

---

## 2. Brief Summary

Built a fully interactive **Microsoft Excel dashboard** that transforms raw, disconnected internet sales data from Adventure Works Cycles into clear business insights — covering sales trends, product performance, customer behaviour, and regional analysis across 5 years and 6 countries.

---

## 3. Overview

**Adventure Works Cycles** is a fictional bicycle manufacturer and a widely used Microsoft sample dataset. This project takes their raw internet sales data — spread across **8 separate Excel files** — and goes through the full data analyst workflow:

```
Raw Files  →  Data Cleaning  →  Data Integration  →  Calculations  →  Dashboard
```

The dashboard is designed to answer **3 core business questions:**

| # | Business Question |
|---|---|
| 1 | 🚲 Which products and categories drive the most revenue and profit? |
| 2 | 👥 Who are the most valuable customers? |
| 3 | 🌍 Which regions and countries perform best? |

The final output is a **single Excel workbook** with a clean, interactive dashboard complete with KPI cards, charts, pivot tables, and slicers — usable by any business stakeholder without technical knowledge.

---

## 4. Problem Statement

The raw data for Adventure Works Cycles had **7 key problems** that made it impossible to analyse directly:

| # | Problem | Impact |
|---|---|---|
| 1 | Sales data split across **2 separate files** covering different years | Could not analyse the full 5-year trend |
| 2 | Sales records only had **ID numbers** (e.g. `ProductKey = 353`) — no readable names | Charts would show meaningless numbers |
| 3 | **No calculated fields** for Sales Amount, Production Cost, or Profit | Could not measure revenue or profitability |
| 4 | Date fields stored as **Excel serial numbers** (e.g. `41274` instead of `31/12/2012`) | Could not filter or group by date |
| 5 | **2 columns were 100% empty** across all 60,398 rows | Wasted space, caused confusion |
| 6 | `Unit price` in DimProduct stored as **Text** with blank spaces instead of numbers | Could not use in any calculation |
| 7 | **No visual summary existed** — data was spread across 8 files with no dashboard | Stakeholders had no way to understand performance |

**Goal:** Clean all 8 files, connect them together using VLOOKUP, calculate the right metrics, and deliver one interactive dashboard that answers the 3 business questions clearly.

---

## 5. Dataset

### Source Files

| File | Type | Rows | What It Contains |
|---|---|---|---|
| `FactInternetSales.xlsx` | Fact Table | 5,627 | Sales transactions 2010–2012 |
| `Fact_Internet_Sales_New.xlsx` | Fact Table | 54,771 | Sales transactions 2013–2014 |
| `Dimcustomer.xlsx` | Dimension | 18,484 | Customer names, demographics, location |
| `DimProduct.xlsx` | Dimension | 606 | Product names, costs, sub-category keys |
| `DimProductCategory.xlsx` | Dimension | 4 | 4 top-level product categories |
| `DimProductSubCategory.xlsx` | Dimension | 37 | 37 product sub-categories |
| `DimSalesterritory.xlsx` | Dimension | 10 | Territories across 3 continents |
| `DimDate.xlsx` | Dimension | 3,652 | Full date calendar with fiscal year fields |

### Combined Dataset at a Glance

| Metric | Value |
|---|---|
| 📦 Total Sales Records | 60,398 |
| 📅 Date Range | 2010 – 2014 (5 years) |
| 👥 Unique Customers | 18,484 |
| 🛍️ Total Products | 606 across 4 categories, 37 sub-categories |
| 🌍 Territories | 10 across 6 countries and 3 continental groups |
| 💰 Total Sales Revenue | $29,358,677 |
| 🏭 Total Production Cost | $17,277,794 |
| 📈 Total Profit | $12,080,884 |
| 📊 Profit Margin | 41.1% |
| 🛒 Total Orders | 27,659 |
| 🎯 Average Order Value | $1,061 |

### How the Tables Connect

```
Sales Data (Fact Table)
   │
   ├── CustomerKey       ──────► DimCustomer
   │                              (Full Name, Gender, Income, Location)
   │
   ├── ProductKey        ──────► DimProduct
   │                              │
   │                              └── ProductSubcategoryKey ──► DimProductSubCategory
   │                                                               │
   │                                                               └── ProductCategoryKey ──► DimProductCategory
   │
   └── SalesTerritoryKey ──────► DimSalesterritory
                                  (Region, Country, Territory Group)
```

---

## 6. Tools and Technologies

| Tool | Purpose |
|---|---|
| **Microsoft Excel** | Primary tool — data cleaning, integration, calculations, and dashboard |
| **VLOOKUP + IFERROR** | Connecting all 8 dimension tables to the main Sales Data fact table |
| **Excel Formulas** | DATE, TEXT, YEAR, MONTH, MOD, TRIM, VALUE, IFERROR, SUM, COUNTA |
| **Pivot Tables** | Aggregating data by Year, Month, Product, Customer, and Region |
| **Pivot Charts** | Bar, Line, Pie, Donut, Clustered, Stacked, and Combo charts |
| **Slicers** | Interactive filter buttons connected to all pivot tables simultaneously |
| **Conditional Formatting** | KPI card highlighting and data bar visualisation |

---

## 7. Methods

### Step 1 — Combine the Two Sales Files

The two fact tables covered different years but had identical columns. They were stacked into one single sheet called `Sales_Data` giving **60,398 total rows** ready for analysis.

---

### Step 2 — Data Cleaning

| Issue Found | Fix Applied |
|---|---|
| `CarrierTrackingNumber` — 100% empty (60,398 blanks) | Deleted the column |
| `CustomerPONumber` — 100% empty (60,398 blanks) | Deleted the column |
| `OrderDate`, `DueDate`, `ShipDate` stored as serial numbers | Formatted cells as Date (Ctrl+1 → Date) |
| `BirthDate`, `DateFirstPurchase` stored as serial numbers | Formatted cells as Date |
| `Title` column — 99.5% empty | Deleted the column |
| `Suffix` column — 100% empty | Deleted the column |
| `NameStyle` column — always `False` | Deleted the column |
| `Unit price` stored as Text with blank spaces | Cleaned with `=VALUE(TRIM())` |
| `Status` — 200 blank cells in DimProduct | Filled with `"Unknown"` via Find & Replace |
| `Color` — 254 blank cells in DimProduct | Filled with `"N/A"` via Find & Replace |
| Row 11 in DimSalesterritory — all values blank | Deleted the invalid row |

---

### Step 3 — Data Integration using VLOOKUP

All lookup tables were connected to `Sales_Data` using VLOOKUP formulas. The formulas must be entered **in this exact order** because each one depends on the previous:

```excel
-- STEP 1: Get ProductSubcategoryKey from Dim Product  →  Col AR
=IFERROR(VLOOKUP(B2,'Dim Product'!$A:$D,4,0),"")

-- STEP 2: Get Sub-Category Name from product Sub-cat  →  Col AS
=IFERROR(VLOOKUP(AR2,'product Sub-cat'!$A:$C,3,0),"N/A")

-- STEP 3: Get ProductCategoryKey from product Sub-cat  →  Col AT
=IFERROR(VLOOKUP(AR2,'product Sub-cat'!$A:$F,6,0),"")

-- STEP 4: Get Category Name from Dim Product Category  →  Col AU
=IFERROR(VLOOKUP(AT2,'Dim Product Category'!$A:$C,3,0),"N/A")

-- STEP 5: Get Territory Group from Dim Sales terri  →  Col AV
=IFERROR(VLOOKUP(I2,'Dim Sales terri'!$A:$E,5,0),"Unknown")
```

---

### Step 4 — Feature Engineering

New calculated columns were added to `Sales_Data`:

```excel
-- Convert OrderDateKey (e.g. 20121231) to a real date
=DATE(LEFT(C2,4), MID(C2,3,2), RIGHT(C2,2))

-- Date breakdown
Year              = YEAR(Date)
Month No          = MONTH(Date)
Month Full Name   = TEXT(Date,"MMMM")
Quarter           = "Q" & ROUNDUP(MONTH(Date)/3,0)
Year Month        = TEXT(Date,"YYYY-MMM")
Weekday No        = WEEKDAY(Date,2)
Weekday Name      = TEXT(Date,"DDDD")

-- Financial calendar (fiscal year starts July)
Financial Month   = MOD(MONTH(Date)-7+12,12)+1
Financial Quarter = "FQ" & ROUNDUP(MOD(MONTH(Date)-7+12,12)/3+0.01,0)

-- Financial metrics
Sales Amount      = UnitPrice * OrderQuantity * (1 - UnitPriceDiscountPct)
Production Cost   = ProductStandardCost * OrderQuantity
Profit            = Sales Amount - Production Cost
```

---

### Step 5 — KPI Calculations

```excel
Total Sales Amount    = SUM(Sales_Data!AM:AM)          -- $29,358,677
Total Production Cost = SUM(Sales_Data!AN:AN)          -- $17,277,794
Total Profit          = SUM(Sales_Data!AO:AO)          -- $12,080,884
Profit Margin %       = Total Profit / Total Sales     -- 41.1%
Total Orders          = COUNTA(Sales_Data!J:J) - 1    -- 27,659
Average Order Value   = Total Sales / Total Orders     -- $1,061
Unique Customers      = COUNTA(UNIQUE(CustomerKey))    -- 18,484
```

---

### Step 6 — Dashboard Charts

| Section | Chart Type | Title |
|---|---|---|
| Product | Horizontal Bar | Top 10 Products by Sales Amount |
| Product | Donut | Sales % by Product Category |
| Product | Clustered Bar | Sales vs Profit by Category |
| Product | Stacked Bar | Sub-Category Sales by Year |
| Customer | Horizontal Bar | Top 10 Customers by Sales |
| Customer | Column Chart | Customer Count by Year |
| Customer | Line Chart | Monthly Sales Trend |
| Region | Donut | Sales Share by Territory Group |
| Region | Column Chart | Sales by Country |
| Region | Multi-Line | Region Sales Trend by Year |
| Trend | Combo Bar + Line | Sales Amount vs Production Cost by Year |
| Monthly | Grouped Bar Chart | Month and Sales — Sales Amount & Production Cost by Month (1–12) |

---

### Step 7 — Slicers

Five interactive slicers were added to the dashboard and connected to **all pivot tables** via Report Connections:

| Slicer | Values |
|---|---|
| Year | 2010, 2011, 2012, 2013, 2014 |
| Financial Quarter | FQ1, FQ2, FQ3, FQ4 |
| Territory Group | North America, Europe, Pacific |
| Product Category | Bikes, Accessories, Clothing, Components |
| Country | USA, Canada, France, Germany, Australia, United Kingdom |

---

## 8. Key Insights

### 💰 Sales Performance
- 📈 **2013 was the peak year** with **$16.35M in sales** — a 180% jump from 2012 ($5.84M)
- 💰 Overall **profit margin is 41.1%** — healthy and sustainable across all 5 years
- 🛒 **27,659 total orders** placed with an average value of **$1,061 per order**
- 📉 2014 shows only $45K — because the dataset captures only the first few weeks of January 2014

### 🚲 Product Insights
- 🥇 **Mountain-200 Black, 46** is the single best-selling product at **$1,373,470**
- 🚵 The entire **top 10 are premium bikes** — Mountain-200 and Road-150 product families
- 🏷️ **Bikes dominate revenue** across all 4 categories — Accessories, Clothing, and Components are secondary
- 🔧 Components sub-category sells at lower price points but could carry higher margins

### 👥 Customer Insights
- 👥 **18,484 unique customers** made at least one purchase over 5 years
- 🥇 Top customer is **Samuel C Mitchell** at **$6,119 in total lifetime sales**
- 📊 The **top 10 customers all spend $6,035–$6,119** — a tight, consistent band of high-value buyers
- 💡 Average customer lifetime value of **$1,589** suggests a loyal, medium-to-high spend customer base

### 🌍 Regional Insights
- 🌎 **North America** is the largest territory group by order volume
- 🇬🇧 **United Kingdom** is the top individual country by regional sales revenue
- 🌏 **Pacific (Australia)** is a significant and growing market
- 🗺️ Business spans **6 countries across 3 continents** — a genuinely global customer base

---

## 9. Dashboard / Output

**File:** `Sales_Dashboard.xlsx`

### Sheet Structure

| Sheet | Purpose |
|---|---|
| `DASHBOARD` | Main interactive dashboard — KPI cards, charts, slicers |
| `Sales_Data` | Combined and cleaned fact table (60,398 rows × 46 columns) |
| `Month and Sales` | Pivot — Monthly sales with Year filter |
| `Year wise Sales` | Pivot — Year-over-year comparison |
| `Month wise Sales` | Pivot — Month-on-month trend |
| `pie chart` | Pivot — Quarterly distribution |
| `combo chart` | Pivot — Sales vs Production Cost |
| `DimCustomer` | Customer dimension table (cleaned) |
| `Dim Product` | Product dimension table (cleaned) |
| `Dim Product Category` | 4 product categories |
| `product Sub-cat` | 37 product sub-categories |
| `Dim Sales terri` | 10 sales territories |

### Dashboard Layout

```
┌───────────────────────────────────────────────────────────────────┐
│  SLICERS:  Year  |  Financial Quarter  |  Territory Group         │
│            Product Category  |  Country                           │
├──────────────┬───────────────┬──────────────────┬─────────────────┤
│  Total       │  Total        │  Total Product   │  Total          │
│  Profit      │  Sales        │  Cost            │  Orders         │
│ 12,080,883   │ 29,358,677    │  17,277,794      │  60,398         │
├──────────────────────────────────┬────────────────────────────────┤
│  📈 Month-Wise Sales (Line)      │  📊 Month and Sales            │
│                                  │     (Grouped Bar — Blue+Orange)│
├──────────────────────────────────┼────────────────────────────────┤
│  📊 Customer Performance (Bar)   │  🍩 Regional Performance (Pie) │
├──────────────┬───────────────────┴────────────────────────────────┤
│  🍩 Quarter- │  📊 Year-Wise Sales │  📊 Product Performance      │
│  wise Sales  │  (Column Chart)     │  (Horizontal Bar)             │
│  (Donut)     │                     │                               │
├──────────────┴─────────────────────┴──────────────────────────────┤
│        📈 Sales Amount vs Production Cost by Year (Combo Chart)   │
└───────────────────────────────────────────────────────────────────┘
```

---

## 10. How to Run This Project

### Requirements
- Microsoft Excel **2016 or later**
- All 8 source `.xlsx` files in the **same folder**

### Instructions

**Step 1 — Combine the two sales files**
```
1. Open FactInternetSales.xlsx and Fact_Internet_Sales_New.xlsx
2. In Fact_Internet_Sales_New → select all rows except the header → Ctrl+C
3. In FactInternetSales → go to the last empty row → Ctrl+V
4. Rename the sheet tab to: Sales_Data
5. File → Save As → Sales_Dashboard.xlsx
```

**Step 2 — Copy all dimension sheets into the workbook**
```
Open each dimension file
→ Right-click the sheet tab → Move or Copy
→ Select Sales_Dashboard.xlsx → tick Create a copy → OK

Repeat for: Dimcustomer, DimProduct, DimSalesterritory,
            DimProductCategory, DimProductSubCategory
```

**Step 3 — Clean the data**
```
In Sales_Data:
  → Delete columns: CarrierTrackingNumber, CustomerPONumber
  → Format as Date: OrderDate, DueDate, ShipDate
  → (Ctrl+1 → Number tab → Date → DD/MM/YYYY → OK)

In DimCustomer:
  → Format as Date: BirthDate, DateFirstPurchase
  → Delete columns: Title, Suffix, NameStyle

In DimProduct:
  → Fix Unit price column: add new col → =VALUE(TRIM(B2)) → fill down
  → Fill blank Status cells with Unknown  (Ctrl+H)
  → Fill blank Color cells with N/A       (Ctrl+H)

In DimSalesterritory:
  → Delete row 11 (the row with all blank region/country values)
```

**Step 4 — Add VLOOKUP columns (in this exact order)**
```
Click Sales_Data → go to cell AR1 → add header → go to AR2 → type formula → fill down
Repeat for each row below:

AR1 = ProductSubcategoryKey   AR2: =IFERROR(VLOOKUP(B2,'Dim Product'!$A:$D,4,0),"")
AS1 = SubcategoryName         AS2: =IFERROR(VLOOKUP(AR2,'product Sub-cat'!$A:$C,3,0),"N/A")
AT1 = ProductCategoryKey      AT2: =IFERROR(VLOOKUP(AR2,'product Sub-cat'!$A:$F,6,0),"")
AU1 = CategoryName            AU2: =IFERROR(VLOOKUP(AT2,'Dim Product Category'!$A:$C,3,0),"N/A")
AV1 = SalesTerritoryGroup     AV2: =IFERROR(VLOOKUP(I2,'Dim Sales terri'!$A:$E,5,0),"Unknown")

To fill down: copy the formula cell → click the cell below → Ctrl+Shift+End → Ctrl+V
```

**Step 5 — Add calculated columns**
```
Add these columns to Sales_Data (fill each down to all 60,398 rows):

Date             =DATE(LEFT(C2,4),MID(C2,3,2),RIGHT(C2,2))
Year             =YEAR(Date_cell)
Month No         =MONTH(Date_cell)
Month Full Name  =TEXT(Date_cell,"MMMM")
Quarter          ="Q"&ROUNDUP(MONTH(Date_cell)/3,0)
Year Month       =TEXT(Date_cell,"YYYY-MMM")
Weekday No       =WEEKDAY(Date_cell,2)
Weekday Name     =TEXT(Date_cell,"DDDD")
Financial Month  =MOD(MONTH(Date_cell)-7+12,12)+1
Fin. Quarter     ="FQ"&ROUNDUP(MOD(MONTH(Date_cell)-7+12,12)/3+0.01,0)
Sales Amount     =N2*M2*(1-P2)
Production Cost  =R2*M2
Profit           =SalesAmount_cell - ProductionCost_cell
```

**Step 6 — Create Pivot Tables and Charts**
```
1. Click any cell inside Sales_Data
2. Insert tab → PivotTable → New Worksheet → OK
3. Build one pivot per chart
4. Insert chart from each pivot (Insert → recommended chart type)
5. Move charts to the DASHBOARD sheet (cut and paste)
```

**Step 7 — Add Slicers**
```
1. Click any pivot table
2. PivotTable Analyze tab → Insert Slicer
3. Select: Year, Quarter, SalesTerritoryGroup, CategoryName, Country
4. Click OK — slicer boxes appear
5. For each slicer: right-click → Report Connections
   → tick ALL pivot tables in the list → OK
   (This makes one slicer control every chart at the same time)
```

**Step 8 — Refresh all data**
```
Press Ctrl + Alt + F5
This refreshes all pivot tables at once
```

---

## 11. Results and Conclusion

### Results Summary

| What Was Done | Outcome |
|---|---|
| Combined 2 sales files | Single `Sales_Data` sheet — 60,398 rows, 46 columns |
| Cleaned 8 files | Removed 2 empty columns, fixed date formats, corrected data types |
| Connected all tables | 5 VLOOKUP formulas linking Product, Customer, Territory, Category |
| Calculated metrics | Sales Amount, Production Cost, Profit, Margin % — all formula-driven |
| Built 11 charts | Product, Customer, and Regional performance fully visualised |
| Created 6 KPI cards | Total Sales, Cost, Profit, Margin %, Orders, Avg Order Value |
| Added 5 slicers | Year, Quarter, Region, Category, Country — all connected to all pivots |
| Delivered dashboard | One clean Excel file — interactive and usable by non-technical stakeholders |

### Conclusion

Adventure Works Cycles operates a **strong and profitable business** with a 41.1% profit margin. Revenue peaked in **2013 at $16.35M**, driven almost entirely by the premium Bikes category — particularly the Mountain-200 and Road-150 product lines. The customer base of **18,484 buyers** spends consistently, with top customers clustered tightly in the $6,000–$6,119 lifetime spend range.

North America is the dominant market by volume, while Europe (especially the United Kingdom) and Pacific (Australia) contribute meaningfully and represent real growth opportunities. The dashboard gives decision-makers a clear, filterable, real-time view of the entire business — making it a practical tool for sales managers, product leads, and executives.

---
## 📁 Project Files

```
Sales_Dashboard.xlsx           ← Main workbook — all sheets, charts, and dashboard
FactInternetSales.xlsx         ← Source data: Sales 2010–2012
Fact_Internet_Sales_New.xlsx   ← Source data: Sales 2013–2014
Dimcustomer.xlsx               ← Lookup: Customer details
DimProduct.xlsx                ← Lookup: Product details
DimProductCategory.xlsx        ← Lookup: 4 product categories
DimProductSubCategory.xlsx     ← Lookup: 37 product sub-categories
DimSalesterritory.xlsx         ← Lookup: Territory and region data
DimDate.xlsx                   ← Lookup: Date calendar
README.md                      ← This file
```

---

> 📌 **Dataset:** Adventure Works Cycles — Microsoft Sample Dataset
> 🛠️ **Tool:** Microsoft Excel
> 📅 **Analysis Period:** 2010 – 2014
> 📁 **Project Type:** End-to-End Data Analytics Project
