# 📊 Regional Revenue Dashboard – 2021

Comprehensive Excel-based analytics dashboard for exploring **regional performance**, **store rankings**, and **product-level revenue insights** across multiple U.S. regions during **July 2021**.

This project includes:

- A **fully automated dashboard**
- Complete **data preparation tables**
- Month-over-month (MoM) and year-over-year (YoY) KPIs
- Store- and product-level performance analytics

![DashboardGif](/Imgs/Dashboard.gif)

---

## 📌 Overview

The dashboard provides a detailed breakdown of **July 2021 performance**, comparing:

- Region-level revenue
- Store-level rankings
- Product performance
- Month-over-month revenue changes
- Year-over-year revenue trends

It is built using:

- **Excel formulas**
- **Dynamic ranges / data validation**
- **Interactive filters**

---

## 🚀 Key Features

### ⭐ **1. KPI Summary**

- **Total Revenue:** $94,953
- **MoM Δ:** +4%
- **YoY Δ:** +34%

![KPIs](/Imgs/KPIs.PNG)

### ⭐ **2. Regional & Store Performance**

Compare how each store performed relative to other regions:

- O'Hare – **19% MoM growth**
- Lincoln Park – **24% MoM growth**
- Millenium – **–13% MoM decline**
- Michigan Ave – **–18% MoM decline**

![Store](/Imgs/Store.PNG)

### ⭐ **3. Revenue Trend Analysis (2020 vs 2021)**

A line chart visualizing:

- Month-by-month performance
- Seasonal patterns
- Differences between years

![Trend](/Imgs/Trend.PNG)

### ⭐ **4. Product Performance**

Highlights the best and worst contributing products:

#### 🟩 **Top Performing Products**

| Product       | Revenue | MoM Δ   |
| ------------- | ------- | ------- |
| Lego Bricks   | $19,035 | +$4,719 |
| Etch A Sketch | $4,408  | +$2,897 |
| Nerf Gun      | $5,617  | +$2,539 |

#### 🟥 **Bottom Performing Products**

| Product        | Revenue | MoM Δ   |
| -------------- | ------- | ------- |
| Magic Sand     | $7,243  | –$3,278 |
| Colorbuds      | $4,272  | –$2,338 |
| Action Figures | $3,966  | –$2,335 |

![Tables](/Imgs/Tables.PNG)

---

## 🛠️ Data Preparation (Data Prep Sheet)

The **Data Prep** sheet includes structured tables for:

### ✔ Region Filters

- Dynamic selection for: Los Angeles, New York, Chicago

![Filters](/Imgs/Filters.PNG)

### ✔ KPIs Table

- Total Revenue
- Last Year Revenue
- Last Month Revenue
- YoY% and MoM% calculations

![KPIsPrep](/Imgs/KPIsPrep.PNG)

### ✔ Revenue Trend Table

- Revenue for each month (2020 vs 2021)
- Current period selection logic

![RevenueTrend](/Imgs/RevenueTrend.PNG)

### ✔ Store Performance

- Revenue by store
- MoM %, YoY %, and rank
- Sorted performance table

![StorePerformance](/Imgs/StorePerformance.PNG)

### ✔ Product Performance

- Revenue, PM Revenue, MoM Δ
- Top / Bottom ranks

This sheet powers the entire dashboard through formulas like:

- `XLOOKUP`
- `INDEX/MATCH`
- `SUMIFS`
- `IFERROR`
- `RANK.EQ`
- Dynamic dropdown lists

**Total Revenue**
=SUMIFS(Data[Revenue],Data[Region],Region,Data[Month],CurrMonth,Data[Year],CurrYear)

**Current Year Revenue for Trend**
=IF(G3>CurrMonth,NA(),SUMIFS(Data[[Revenue]:[Revenue]],Data[[Region]:[Region]],Region,Data[[Month]:[Month]],'Data Prep'!$G3,Data[[Year]:[Year]],'Data Prep'!J$2))

**Lookup Store Region**
=VLOOKUP(M3,Data[[Store Name]:[Region]],2,0)

**Store Ranking**
=RANK.AVG(O3,$O$3:$O$12)

**Get Store After Sorting (Top Stores)**
=INDEX($M$3:$R$12,MATCH($V3,$R$3:$R$12,0),MATCH(W$2,$M$2:$R$2,0))

**Previous Month Revenue for Every Product**
=SUMIFS(Data[Revenue],Data[Region],Region,Data[Month],PrevMonth,Data[Year],PMYear,Data[Product Name],'Data Prep'!AA3)

---

## 📸 Dashboard Preview

![Dashboard](/Imgs/Regional%20Revenue%20Dashboard.PNG)
