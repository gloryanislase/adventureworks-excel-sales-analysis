# 📊 Adventure Works Sales Analysis — Advanced Excel Dashboard

[![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)](#)
[![Data Analysis](https://img.shields.io/badge/Data_Analysis-005571?style=for-the-badge)](#)
[![Dashboard](https://img.shields.io/badge/Dashboard-0F6CBD?style=for-the-badge)](#)

An Excel-based **sales analytics and dashboarding project** using the Adventure Works dataset to analyze revenue, profitability, product performance, customer contribution, and time-based trends. The project demonstrates practical **Advanced Excel skills** through formula-driven analysis, lookup and ranking functions, helper tables, dynamic KPI calculations, and interactive dashboard design. The goal is to transform transactional data into concise business insights using Microsoft Excel.

## 📸 Dashboard Preview

### Time Series Dashboard

![Time Series Dashboard](./assets/time_series_dashboard_overview.png)

### Product & Customer Analysis

![Product & Customer Dashboard](./assets/detail_dashboard_by_product_&_customer.png)

---

## 🎥 Dashboard Demo  
Watch a short walkthrough of the interactive Excel dashboard, including filtering, dynamic metric selection, dashboard navigation, and KPI updates.  
[▶️ Watch Dashboard Demo](https://youtu.be/5O523CglT04)

---

## 📌 About This Project
**Objective**   
Build an end-to-end Excel sales dashboard that's functionally on par with a Power BI dashboard — complete with a relational data model, DAX calculations, and multi-page navigation — to show that Excel can be a serious BI tool, not just a static spreadsheet.

**Methodology**  
Raw AdventureWorksDW data was cleaned with **Power Query** (one query per table), loaded into the **Data Model (Power Pivot)**, and connected in a **star schema** (`FactInternetSales` as the fact table, with 5 dimension tables around it). Every core metric is computed as a **DAX measure** (not a static column) so it recalculates automatically with the active slicer selection. Visuals are built from a combination of **PivotTable + PivotChart + Slicer**, restyled to resemble a modern dashboard layout. Cross-dashboard navigation and the Clear Filter button run on a **VBA macro** that resets every `SlicerCache`.

---

## 📂 Repository Structure

```text
adventureworks-excel-sales-analysis/
│
├── README.md
├── LICENSE
│
├── workbook/
│   └── Dashboard Master Class_AdventureWorksSales.xlsm
│
└── assets/
    ├── time_series_dashboard.png
    └── detail_dashboard.png
```

---
## 📊 Key Findings

### 1. Strong revenue and profit growth

Revenue increased from approximately **$33.4M in 2005** to **$101.9M in 2008**, while profit increased from **$13.4M** to **$42.2M**. Profit margin remained relatively stable at around **40–42%** throughout the period.

### 2. Q4 dominates profitability

**Q4 contributed approximately 41.4% of total profit**, making it the strongest quarter. December was the highest-profit month, contributing approximately **$7.57M**.

### 3. Profit is concentrated among a small group of products

The **Top-5 products generated approximately 34.9% of total profit**, with Mountain-200 variants among the leading contributors.

### 4. Geographic contribution is concentrated

**Australia and the United States contributed approximately 60.1% of total profit combined**, indicating a strong concentration of profitability across these markets.

### 5. Customer profitability is broadly distributed

The **Top-5 customers contributed only around 0.5% of total profit**, suggesting that overall profitability is not heavily dependent on a small number of individual customers.

## 🧩 Advanced Excel Skills

| Area                        | Techniques                                                    |
| --------------------------- | ------------------------------------------------------------- |
| **Lookup & Reference**      | `VLOOKUP`, `INDEX`, `MATCH`                                   |
| **Conditional Logic**       | `IF`, `IFS`, `IFERROR`                                        |
| **Ranking & Aggregation**   | `LARGE`, `SUM`, `AVERAGE`, `COUNTA`                           |
| **Analytical Calculations** | YoY comparison, profit contribution, margins, ratios          |
| **Dashboarding**            | KPI cards, charts, filters, dynamic selections                |
| **Excel Modeling**          | Helper tables, named ranges, formula-driven calculation layer |

The workbook separates analytical calculations from dashboard presentation, allowing the visual components to update dynamically based on user selections.

## 🎛️ Dashboard & Analysis

The workbook combines several analytical views:

* **Time Series Analysis:** revenue, profit, COGS, transactions, order quantity, monthly and quarterly trends, and year-over-year comparisons.
* **Product Analysis:** top profitable products, profit concentration, product color, and price-group performance.
* **Customer Analysis:** top profitable customers, age-group contribution, gender, and country-level profitability.
* **Interactive Dashboarding:** dynamic KPI reporting with year, month, and country selections.

## 💡 Business Insights

The analysis highlights several patterns that can support further business investigation:

* Profitability is strongly concentrated toward the end of the year.
* A relatively small number of products contributes a substantial share of total profit.
* Higher-priced products account for the majority of observed profit.
* Overall customer profitability is broadly distributed rather than dependent on a few customers.
* Profit contribution varies considerably across countries.

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* Advanced Excel Formulas
* Lookup & Reference Functions
* Data Analysis
* Dashboarding & Data Visualization



## 🚀 How to Use

1. Download the Excel workbook from the `workbook/` folder.
2. Open it using **Microsoft Excel Desktop**.
3. Explore the dashboards and use the available filters to interact with the analysis.
4. Review the supporting analysis sheets to see how the dashboard metrics are calculated.

> **Dataset note:** Adventure Works is used as a sample dataset for demonstrating analytical and Advanced Excel capabilities. The project is intended as a portfolio case study rather than a representation of current market conditions.

## 📬 Contact

Open to opportunities related to **Data Analyst, Business Intelligence, and Analytics** roles.

* **LinkedIn:** [linkedin.com/in/gloryanisveronicalase](https://linkedin.com/in/gloryanisveronicalase)
* **Email:** [gloryanislase@gmail.com](mailto:gloryanislase@gmail.com)
