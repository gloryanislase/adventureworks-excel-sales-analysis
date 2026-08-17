# 📊 Adventure Works Sales Analysis — Advanced Excel Dashboard

[![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)](#)
[![Data Analysis](https://img.shields.io/badge/Data_Analysis-005571?style=for-the-badge)](#)
[![Dashboard](https://img.shields.io/badge/Dashboard-0F6CBD?style=for-the-badge)](#)

An Excel-based **sales analytics and dashboarding project** using the Adventure Works dataset to analyze revenue, profitability, product performance, customer contribution, and time-based trends. The project demonstrates practical **Advanced Excel skills** through formula-driven analysis, lookup and ranking functions, helper tables, dynamic KPI calculations, and interactive dashboard design. The goal is to transform transactional data into concise business insights using Microsoft Excel.

---

## 📌 About This Project
**Objective**   
Build an end-to-end Excel sales dashboard that's functionally on par with a Power BI dashboard — complete with a relational data model, DAX calculations, and multi-page navigation — to show that Excel can be a serious BI tool, not just a static spreadsheet.

**Methodology**  
Raw AdventureWorksDW data was cleaned with **Power Query** (one query per table), loaded into the **Data Model (Power Pivot)**, and connected in a **star schema** (`FactInternetSales` as the fact table, with 5 dimension tables around it). Every core metric is computed as a **DAX measure** (not a static column) so it recalculates automatically with the active slicer selection. Visuals are built from a combination of **PivotTable + PivotChart + Slicer**, restyled to resemble a modern dashboard layout. Cross-dashboard navigation and the Clear Filter button run on a **VBA macro** that resets every `SlicerCache`.

---

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

## 📊 Key Findings

### 1. Strong revenue and profit growth  
Revenue increased from **$33.37M in 2005** to **$101.86M in 2008**, while profit increased from **$13.40M** to **$42.16M**.

### 2. Q4 dominates profitability

**Q4 contributed approximately 41.4% of total profit**, with **December generating $7.57M**, the highest monthly profit in the analysis.

### 3. Profit is concentrated among a small group of products

The **Top-5 products generated approximately 34.9% of total profit**, led by Mountain-200 variants.

### 4. Weekday sales contribute the majority of profit

**Weekdays contributed 72% of total profit**, with **Wednesday–Friday accounting for 43.8%**, indicating that profitability is concentrated toward the middle and end of the working week.

### 5. Customer profitability is broadly distributed

The **Top-5 customers contributed only around 0.5% of total profit**, indicating low dependence on a small number of individual customers.

---

## 🧩 Advanced Excel & BI Techniques

| Category | Tools |
| --- | --- |
| **Data Preparation** | Power Query |
| **Data Modeling** | Power Pivot (Data Model), star schema |
| **Calculations** | DAX (CALCULATE, SUMX, FILTER, DIVIDE, etc.) |
| **Visualization** | PivotTable, PivotChart, excel chart |
| **Interactivity** | Slicers, Metric Selector, Dashboard Navigation |
| **Automation** | VBA Macro (navigation & Clear Filter) |

### Data Model

The workbook uses a **star schema** with `FactInternetSales` as the central fact table connected to:

- `DimDate`
- `DimProduct`
- `DimCustomer`
- `DimGeography`
- `DimSalesTerritory`

This structure allows transaction-level sales to be analyzed across time, products, customers, and geographic dimensions.

---

## 📊 Dashboard Pages

### Time Series Dashboard  
Provides an overview of:  
- Revenue, Profit, COGS, Margin, Transactions, and Quantity
- Year-level performance
- Monthly profit
- Weekday vs. weekend contribution
- Quarterly profit distribution
- Dynamic Revenue / Profit / Transaction metric selection

### Product & Customer Dashboard  
Provides:  
- Top-5 profitable products and customers
- Product availability analysis
- Profit by product color
- Price-group analysis
- Profit contribution by age and gender
- Country-level profit contribution

### Supporting Analysis  
The workbook also includes supporting analysis pages for:  
- Time series analysis
- Product analysis
- Customer analysis

These pages provide more detailed PivotTable/PivotChart views behind the main dashboards.

---

## 🎛️ Interactive Features

- **Global Slicers:** Filter by Country, Year, and Month.
- **Dynamic Metric Selection:** Switch between Revenue, Profit, and Transactions without creating separate charts.
- **Cross-Filtering:** Selecting a category or data point updates related calculations and contribution metrics.
- **Dashboard Navigation:** VBA buttons allow users to move between dashboard pages.
- **Clear Filter:** VBA resets the active slicer selections with one click.

---

## 🚀 How to Use

1. Download the Excel workbook from the `workbook/` folder.
2. Open it using **Microsoft Excel Desktop**.
3. Explore the dashboards and use the available filters to interact with the analysis.
4. Review the supporting analysis sheets to see how the dashboard metrics are calculated.

> **Dataset note:** Adventure Works is used as a sample dataset for demonstrating analytical and Advanced Excel capabilities. The project is intended as a portfolio case study rather than a representation of current market conditions.

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

## 📬 Contact

Open to opportunities related to **Data Analyst, Business Intelligence, and Analytics** roles.

* **LinkedIn:** [linkedin.com/in/gloryanisveronicalase](https://linkedin.com/in/gloryanisveronicalase)
* **Email:** [gloryanislase@gmail.com](mailto:gloryanislase@gmail.com)
