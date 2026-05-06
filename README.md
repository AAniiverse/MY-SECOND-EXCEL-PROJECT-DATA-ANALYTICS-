# MY-SECOND-EXCEL-PROJECT-DATA-ANALYTICS-
This project shows complete analysis of the capstone adidas dataset(Given not extracted).


# Adidas Sales Data Analysis — Capstone Project

**Tool:** Microsoft Excel  
**Dataset Source:** External Learning Source   
**Raw Format:** Excel Workbook (.xlsx)  
**Analysis Type:** Full 360° Sales Performance Audit  

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Business Questions](#2-business-questions)
3. [Dataset Description](#3-dataset-description)
4. [Data Collection](#4-data-collection)
5. [Data Cleaning](#5-data-cleaning)
6. [Data Analysis](#6-data-analysis)
7. [Dashboard Design](#7-dashboard-design)
8. [Key Findings](#8-key-findings)
9. [Strategic Recommendations](#9-strategic-recommendations)
10. [Lessons Learnt](#10-lessons-learnt)
11. [Tools & Skills Used](#11-tools--skills-used)
12. [About This Project](#12-About-This-Project)


---

## 1. Project Overview

This project showcases a full end-to-end analysis of Adidas sales data — from raw dataset ingestion through cleaning, structured analysis, dashboard development, and final strategic interpretation. 

The core objective was to deliver a business-ready performance audit that any Adidas stakeholder could act on. The analysis covers four dimensions simultaneously: **products**, **sales channels**, **geographic regions**, and **retail partners** — with every insights grounded directly in the data.

The final output is a set of Excel dashboards showing key details at a glance for necessary insights and decisions.

---

## 2. Business Questions

Before touching the data, the following questions were defined to frame the entire analysis:

| # | Business Question |
|---|---|
| 1 | Which product categories generate the most revenue and units sold? |
| 2 | Which sales channel (Online, Outlet, In-Store) is most efficient by revenue and profit margin? |
| 3 | Which geographic region is the strongest performer, and why might others be lagging? |
| 4 | Which retail partners are delivering the most value, and which are underperforming relative to their potential? |
| 5 | What is the monthly sales trend, and when is the peak demand period? |
| 6 | What are the key risks to the current revenue structure? |
| 7 | What concrete actions should Adidas leadership prioritise? |

6.	Are there specific products that show strong growth potential or declining demand?
7.	How do Adidas sales compare to projected targets or benchmarks in different markets?
8.	What are the top-performing product lines across regions?


Defining these questions first prevented scope creep and kept every cleaning step, pivot table, and chart directly tied to a decision-making need.

---

## 3. Dataset Description

The dataset was provided as a single Excel workbook (.xlsx). It contained transactional-level Adidas sales records covering a full year across the United States market.

**Columns in the raw dataset:**

| Column | Description |
|---|---|
| Retailer | Name of the retail partner (e.g. West Gear, Foot Locker, Amazon) |
| Retailer ID | Unique identifier for each retailer |
| Invoice Date | Date of the sales transaction |
| Region | Geographic region of the sale (West, Northeast, Southeast, South, Midwest) |
| State | U.S. state of the transaction |
| City | City of the transaction |
| Product | Full product name (e.g. "Men's Street Footwear") |
| Price per Unit | Selling price per individual unit |
| Units Sold | Quantity of units sold in that transaction |
| Total Sales | Pre-filled total sales figure (found to be incorrect — see Section 5) |
| Operating Profit | Pre-filled profit figure (found to be incorrect — see Section 5) |
| Operating Margin | Profit margin percentage |
| Sales Method | Sales channel: Online, Outlet, or In-Store |


---

## 4. Data Collection

The dataset was received directly from an external source as a downloadable file. No web scraping, API calls, or external data merging was required at this stage.

Upon first opening the file, two immediate preliminary checks were performed before any formal cleaning began:

1. **Row and column count** — confirmed the full extent of the dataset to understand its size and flag any obvious truncation.
2. **Visual scan of each column** — a quick scroll through every column to spot anything obviously wrong: blank cells, strange values, or formatting inconsistencies visible to the eye before any formula was applied.

This preliminary review is what surfaced the first and most significant data quality issue: the **Total Sales** and **Operating Profit** columns were visually inconsistent with the other figures in the dataset, which led to the deeper investigation described in Section 5.

---

## 5. Data Cleaning

Data cleaning was the most technically demanding phase of this project. The dataset appeared clean on the surface but contained structural errors that would have produced completely incorrect analysis results if left unaddressed.

### 5.1 The Total Sales Inflation Problem

The `Total Sales` column was pre-filled in the dataset. Upon inspection, the values did not match a manual calculation of `Units Sold × Price per Unit`. The discrepancy was not random — every single Total Sales figure was **inflated by exactly 10×** relative to what the arithmetic should produce.

<img width="1567" height="443" alt="Screenshot 2026-05-06 080633" src="https://github.com/user-attachments/assets/1c5e6c53-7f9a-494a-876e-46406247e6e9" />


**Resolution:**
- Created a new calculated column using the formula: `=Units Sold * Price per Unit`
- Replaced the pre-filled `Total Sales` column with this recalculated version
- Verified a sample of rows manually to confirm the new figures were correct

The same 10× inflation was present in the `Operating Profit` column. This was also recalculated from first principles rather than trusted as-is.

> **Why this matters:** Had the inflated figures been used, the total revenue figure would have read approximately $1.2 billion instead of $120 million. Every chart, percentage, and recommendation would have been built on false data.

### 5.2 Product Name Standardisation

The `Product` column contained full descriptive product names that were long and consuming much space which was not good especially for dashboarding. To make analysis cleaner and pivot tables more readable, I...:

- Created a **new helper column** to hold shortened product names using the product name Initials.
- Built a **product bucket list** (a lookup reference table) on the Product Analysis sheet mapping each full product name to its standardised short form

  <img width="429" height="147" alt="Screenshot 2026-05-06 180142" src="https://github.com/user-attachments/assets/db55f3d2-2f9d-4278-919e-525dd481badb" />
  
<p></p>

- Used `XLOOKUP` to populate the new column by returning the correct short form for each row based on the full product name. (VLOOKUP() does same work but XLOOKUP() has been more flexible and high level since its launch, perhaps, it's facing out V- and H-lookup())

This created a consistent, reusable reference that made all downstream pivot tables and charts display clean, uniform category labels.

<img width="1857" height="468" alt="Screenshot 2026-05-06 080258" src="https://github.com/user-attachments/assets/ca4acd2e-22d5-4ec7-904b-e1cc2c6311a4" />


### 5.3 Date Column Verification

The `Invoice Date` column appeared correctly formatted visually. However, a known issue with Excel is that dates can be stored as text strings that look like dates but behave differently in formulas and pivot tables.

**Steps taken:**
- Applied Excel date functions (`YEAR()`, `MONTH()`, `DAY()`) to each date value to confirm they were returning numeric outputs, not errors
- Extracted month numbers separately into a helper column to power the monthly trend analysis
- Verified that no dates fell outside the expected full-year range

> **Principle applied:** A cleaned date column is not actually clean unless *you* have cleaned it yourself. Visual appearance is not a substitute for formula-based verification.

### 5.4 Full Column-by-Column Consistency Check

Every column in the dataset was passed through Excel's `UNIQUE()` function. This returns a deduplicated list of every distinct value in a column, making it immediately visible if:

- A retailer name was spelled two different ways (e.g. "West Gear" vs "WestGear")
- A region had an extra space or capitalisation variant
- A sales method had an unexpected value outside the three expected categories

This step found no additional inconsistencies beyond those already addressed — but it is a non-negotiable verification step. Skipping it creates the risk of silent errors where a mistyped category value gets counted separately in pivot tables and creates phantom rows in analysis.

### 5.5 Cleaning Summary

| Issue Found | Column Affected | Resolution |
|---|---|---|
| Total Sales inflated 10× | Total Sales | Recalculated as `Units Sold × Price per Unit` |
| Operating Profit inflated 10× | Operating Profit | Recalculated from corrected Total Sales |
| Long product names | Product | XLOOKUP against product bucket list |
| Date format unverified | Invoice Date | Confirmed with YEAR/MONTH/DAY functions |
| Potential category inconsistencies | All categorical columns | UNIQUE() scan — no additional issues found |

---
## 6. Data Analysis

With a clean, verified dataset, analysis was conducted using **Excel PivotTables**, **PivotCharts**, and calculated fields. Each business question from Section 2 was answered through a dedicated analysis view.

### 6.1 Product Analysis

**Approach:** Pivot table grouping `Total Sales` and `Units Sold` by `Product` (using the standardised short names from the XLOOKUP column).

**What was measured:**
- Total revenue per product category
- Total units sold per product category
- Implied average revenue per unit (revenue ÷ units — a proxy for pricing strength)

**Finding:** Men's Street Footwear(M.S.F) led on both revenue ($28M) and units (593K). Men's Apparel was the structural underperformer on both dimensions, flagged for corrective action.

<img width="535" height="183" alt="Screenshot 2026-05-06 095334" src="https://github.com/user-attachments/assets/9a43e1a7-35cf-40e8-a516-146cddb270f1" /> 
<img width="456" height="278" alt="Screenshot 2026-05-06 095639" src="https://github.com/user-attachments/assets/91d991c5-c5dd-4d5f-9d88-85d8b858052c" />


### 6.2 Channel Analysis

**Approach:** Pivot table grouping `Total Sales` and `Operating Profit` by `Sales Method`.

**What was measured:**
- Revenue per channel
- Profit per channel
- Profit contribution as a percentage of total profit (not just revenue share)

**Finding:** Online generated the highest revenue ($44.97M) and the highest profit contribution (41%) — making it the most efficient channel. In-Store had the lowest profit contribution (27%) relative to its revenue share (29.7%), indicating high fixed cost drag.

<img width="1302" height="156" alt="Screenshot 2026-05-06 180911" src="https://github.com/user-attachments/assets/3d3589b6-ff31-4fa2-a4a0-a16273bd1e91" />

<p></p>

### 6.3 Regional Analysis

**Approach:** Pivot table grouping `Total Sales` by `Region`.

**What was measured:**
- Revenue by region
- Gap analysis between the highest and lowest performing regions

**Finding:** West led at $36.4M. Midwest trailed at $16.7M — a $19.7M gap that represents either a market penetration failure or an untapped opportunity depending on the growth strategy applied.

<img width="423" height="293" alt="Screenshot 2026-05-06 181227" src="https://github.com/user-attachments/assets/720a4d3f-659a-449f-a342-7f9625fcf29b" />

<p></p>

### 6.4 Retailer Analysis

**Approach:** Pivot table grouping `Total Sales` by `Retailer`.

**What was measured:**
- Revenue per retail partner
- Partner share of total revenue
- Identification of concentration risk (single-partner dependency)

**Finding:** West Gear contributed 27% of total retail revenue ($32M). Amazon contributed only 8% ($10M) — a critical underperformance for the world's largest e-commerce marketplace.

<img width="400" height="273" alt="Screenshot 2026-05-06 181450" src="https://github.com/user-attachments/assets/49b052d3-e8bd-4fc1-a1ae-664c647123c3" />

<p></p>

### 6.5 Monthly Trend Analysis

**Approach:** Pivot table grouping `Total Sales` by extracted `Month` number (derived during cleaning).

**What was measured:**
- Revenue by month across the full year
- Identification of peak and trough periods

**Finding:** July and August are the clear revenue peak — the summer season drives the highest volume. Q4 (October-December) is the softest period of the year, a finding that directly informs inventory planning and promotional strategy.


<img width="413" height="250" alt="Screenshot 2026-05-06 101307" src="https://github.com/user-attachments/assets/a6bc4f91-cebe-4ec4-bc0c-989dca139442" />


---

## 7. Dashboard Design

### 7.1 Tool Choice: Microsoft Excel

Excel was chosen as the sole tool for this project as it was the focus course of study

### 7.2 Dashboard Structure

Three separate dashboard sheets were built:


### **Sales or General Overview** 
High-level KPIs: total revenue, total profit, margin %, total units sold, Revenue by Product, Retailer, Region and Channel, etc

<img width="1598" height="844" alt="Screenshot 2026-05-06 111230" src="https://github.com/user-attachments/assets/fa00583f-18c6-4a11-81a5-9288c672ba05" />

<p></p>


### **Product Analysis** 
Product revenue comparison, breakdown by channel, units by product, product trend, etc.

<img width="1603" height="841" alt="Screenshot 2026-05-06 111253" src="https://github.com/user-attachments/assets/0052a30e-3e43-4d71-8842-4dc7d634d5ad" />

<p></p>

### **Channel Dashboard** 
Channel revenue map; Channel performance comparison and strength in each region, product, retailer; Channel Trend, etc. 

<img width="1599" height="841" alt="Screenshot 2026-05-06 111317" src="https://github.com/user-attachments/assets/f2225b2b-a5fd-4704-a0df-c6ad92e5a30d" />


### 7.3 Colour Choices and Visual Language

Every colour decision in the dashboards was deliberate and functional:

**Primary palette — dark navy/charcoal backgrounds with white text:**  
Chosen to align with Adidas' brand colours and to create strong contrast and reduce eye strain during extended viewing. Dark backgrounds also make data visualisations (charts, numbers) stand out more clearly than light-mode alternatives.

**Green to highlight Top Performers:**
Top performing categories are clearly hightlighted to direct the eyes exactly where it ought to be on the dashboard. Green represent growth as can be pictured in nature, hence, the use to indicate categories that are growing and outperforming others on the list.

**Red (#FF5050 equivalent) to highlight or indicate underperformers:**  
Applied specifically to Men's Apparel revenue, Amazon retailer revenue, Midwest regional figures, and In-Store profit contribution. Red is used sparingly and consistently — it always signals the same thing: a metric that is below what the data suggests it should be.

**Consistent bar chart orientation:**  
All bar charts in the dashboard use horizontal bars for category comparisons (products, retailers, regions) and vertical bars for time-series data (monthly or sales and product trend). This convention makes it immediately clear whether a chart is showing "which category" or "when."

### 7.4 Chart Type Choices

| Analysis | Chart Type | Reason |
|---|---|---|
| Product revenue comparison | Horizontal bar chart | Allows easy reading of long category names without rotation |
| Monthly sales trend | Vertical bar chart | Time flows left to right — vertical bars reinforce temporal progression |
| Channel comparison | Column/card KPIs | Three values are better shown as cards than a chart, differentiated by the highlight colours |
| Regional breakdown | Horizontal bar chart | Five regions with clear rank ordering |
| Retailer performance | Horizontal bar chart | Six retailers — horizontal labels are more readable, and explanatory |

---

## 8. Key Findings

| # | Finding | Data Point |
|---|---|---|
| 1 | Total revenue for the full year | **$120M** |
| 2 | Total operating profit | **$47M** |
| 3 | Overall profit margin | **39%** |
| 4 | Total units sold | **2.48M** |
| 5 | o	🟢Top product by revenue and units | **Men's Street Footwear — $28M, 593K units** |
| 6 | o	🔴Weakest product | **Men's Apparel — lowest revenue and units across all regions** |
| 7 | o	🟢Most efficient channel | **Online — $44.97M revenue, 41% of total profit** |
| 8 | o	🔴Least efficient channel | **In-Store — 29.7% revenue share, only 27% profit contribution** |
| 9 | o	🟢Strongest region | **West — $36.4M** |
| 10 | o	🔴Weakest region | **Midwest — $16.7M (gap of $19.7M vs West)** |
| 11 | o	🟢Top retailer | **West Gear — $32M (27% of retail revenue)** |
| 12 | o	🔴Most underperforming retailer | **Amazon — $10M (8% of retail revenue)** |
| 13 | o	🟢Revenue peak months | **July and August (summer spike)** |
| 14 | o	🔴Softest period | **Q2&Q4 — April-June and October-December** |

---

## 9. Strategic Recommendations

Five recommendation pillars were developed directly from the analysis findings:

### Pillar 1 — Product Strategy
- **Invest behind Men's Street Footwear** — it is the volume and margin leader across all regions. Expand colourways, exclusive partnerships, and dedicated marketing.
- **Restructure Men's Apparel** — reposition with a fresh campaign or rationalise the SKU count. Continued underperformance will cede market share to competitors.
- **Grow Women's Athletic Footwear** — strong trajectory with room to close the gap on Men's. Growing category with demonstrable untapped demand.
- **Overall** — Product categories are stable, not extreme, which is good, meanwhile, lean into the top three products driving revenue - Men’s Street Footwear, Women’s Apparel, and Men’s Athletic Footwear. Reposition weaker categories.

### Pillar 2 — Channel Optimisation
- **Redirect 10–15% of In-Store budget to Online** — Online delivers 41% profit at lower operational cost. Reducing In-Store fixed cost directly improves overall margin.
- **Fix the Amazon channel** — $10M from the world's largest marketplace is a structural failure. Invest immediately in advertising, listing quality, and placement.
- **Use Outlet strategically in Q2&Q4** — the softest revenue period (Apr-Jun & Oct-Dec) is the right window for planned clearance through Outlet to protect margin while moving aged inventory.
- **Improve** — In-store experience and Omnichannel integration (click & collect) - Use in-stores as brand experience hubs, not just sales points.

### Pillar 3 — Regional Focus
- **Study and replicate the West region playbook** — $36.4M did not happen by accident. Extract its retailer mix, product focus, and channel split, then apply those levers to underperforming markets.
- **Run a dedicated Midwest campaign** — at $16.7M it is the weakest market. Increased Sports Direct and Foot Locker presence are the most immediate levers.
- **Accelerate Northeast investment** — $25.1M with clear headroom. Online investment and additional West Gear-style partnerships can move this market materially.
- **Investigate** — Instore and Outlet activities specifically in the south and southeast region respectively. There is a massive concern as they contribute only about 1.65% and 9.69% to a total revenue of over $20M and $21M respectively in both region.

### Pillar 4 — Demand Forecasting
- **Build summer inventory by April** — the July-August spike is consistent and predictable. Stockouts during peak summer are a direct revenue loss. Procurement must start no later than April. Replicate same thing for December-January low peak period starting from September-October
- **Counter Q4 decline proactively** — launch Online-targeted promotions in October-December to counteract the post-summer revenue drop. Use the highest-efficiency channel to smooth the curve.
- **Sustain Men's Footwear safety stock year-round** — this category moves units in every month and every region. Consistent availability is a competitive advantage, and non-negotiable.

### Pillar 5 — Retailer Management
- **Deepen the West Gear partnership** — contractual commitments, co-marketing, and exclusive product drops. Losing this partner removes $32M overnight — treat it as a dependency to protect.
- **Set Amazon a 6-month performance target** — define KPIs, invest in the channel, measure results. If performance does not improve materially, reassess the investment level.
- **Review Walmart immediately** — $11M from Walmart's national footprint should be 2-3× higher. Pricing, placement strategy, and in-store presentation all need examination.
- **Grow Kohl's with Women's Apparel** — demographic alignment makes this a natural pairing. Targeted promotions could unlock meaningful growth from a stable $14M base.

---

## 10. Lessons Learnt

### 10.1 During Data Collection
> **Trust nothing until you have verified it yourself.**

The dataset came from a credible source and looked correct at first glance. But the 10× inflation error in the Total Sales and Operating Profit columns was invisible to the eye — it only became apparent when the numbers were tested against an independent calculation. The lesson: always sanity-check pre-filled columns against their source components (in this case, `Units Sold × Price per Unit`) before treating them as usable data.

### 10.2 During Data Cleaning
> **UNIQUE() is one of the most underrated cleaning tools in Excel.**

Running `UNIQUE()` on every categorical column costs almost no time and eliminates an entire class of silent errors — duplicate values caused by spelling differences, extra spaces, or capitalisation mismatches. A pivot table will count "West Gear" and "west gear" as two separate retailers. `UNIQUE()` catches that before it becomes a problem in your analysis.

> **Date columns that look clean are not necessarily clean.**

Dates stored as text strings are one of the most common silent errors in Excel. They display identically to real date values but break every time-based formula and pivot table grouping. The habit of always verifying dates with `YEAR()`, `MONTH()`, and `DAY()` functions prevents this entirely.

> **XLOOKUP is significantly more versatile than VLOOKUP for this type of standardisation work.**

Using XLOOKUP to map long product names to standardised short forms was far cleaner than manually finding and replacing values — it is auditable (you can see exactly what the formula is doing), reversible (the original column is untouched), and reusable (adding a new product only requires updating the lookup table, not re-cleaning the column).

### 10.3 During Analysis
> **Define your questions before you open a pivot table.**

Starting analysis without predefined business questions leads to exploring data in circles. The seven questions defined at the outset (Section 2) meant every pivot table had a specific purpose — there was no ambiguity about what was being measured or why.

> **Profit contribution percentage reveals more than revenue alone.**

Revenue share tells you how big a channel or product is. Profit contribution percentage tells you how *efficient* it is. The Online channel's 41% profit contribution from 37.4% revenue share is the single most actionable insight in the dataset — and it would have been invisible if the analysis only looked at revenue.

### 10.4 During Dashboard Design
> **Colour is a signal, not a decoration.**

Using red sparingly and consistently — only for genuine underperformers — meant that when a stakeholder saw red on the dashboard, they knew immediately that something required attention. Overusing colour decoratively removes its communicative power entirely.

> **Build for the reader, not for yourself.**

The person reading a dashboard did not spend weeks inside the data. Every chart needs a clear title that states the conclusion, not just the variable name. "Men's Street Footwear Leads All Products by Revenue and Volume" communicates more than "Product Revenue Chart."

### 10.5 During Interpretation
> **Data does not make decisions — analysts do.**

The numbers showed that Amazon generated $10M. The data alone does not say whether that is acceptable or a failure. The interpretation — that it is a critical underperformance for the world's largest marketplace — required understanding the context of what Amazon *should* be capable of delivering. Analytical skill is knowing when a number is a problem, not just knowing what the number is.

---

## 11. Tools & Skills Used

| Category | Detail |
|---|---|
| **Primary Tool** | Microsoft Excel 2021 |
| **Data Cleaning** | Manual formula auditing, `UNIQUE()`, `XLOOKUP()`, `YEAR()`, `MONTH()`, `DAY()`, etc. |
| **Analysis** | PivotTables, PivotCharts, calculated fields |
| **Visualisation** | Excel dashboard sheets, bar charts, KPI cards, colours, fonts, etc. |
| **Core Skills** | Data cleaning, exploratory data analysis, business interpretation, dashboard design(Graphics & Creative design), strategic communication, analytical & critical thinking |

---

## 12. About This Project

**Programme:** Data Analytics Learning Journey  
**Analyst:** Nyong Asuabiat  
**Period:** April.May 2026  
**Dataset:** Adidas U.S. Sales — Full Year  
**Contact:** [LinkedIn](https://linkedin.com/in/aanyy)

---


*This project was completed as part of SelfLearning Journey in Data Analytics. Dataset was gotten from an external learning source. All analysis, interpretations, and recommendations are the work of the analyst and are based solely on the dataset provided.*
