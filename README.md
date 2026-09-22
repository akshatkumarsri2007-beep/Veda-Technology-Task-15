
# Task 15 — Product Count Analysis

**Veda Technology Internship — Data Analytics Track**
**Level 1 · Day 15 of 45**

##  Overview

This task analyzes a Superstore-style dataset to count products by category and identify the category with the largest number of products.

##  Objective

Practice using `COUNTIF` / `COUNTIFS` functions in Excel for categorical data summarization.

## Tools Used

- MS Excel
  - `COUNTIF()`
  - `INDEX()` + `MATCH()`
  - PivotTable (for cross-verification)

##  Files

| File | Description |
|---|---|
| `Superstore_Practice_Task15.xlsx` | Sample dataset + working Excel file with formulas |
| `Task15_Product_Count_Analysis_Report.pdf` | Final task report |

##  Approach

1. Identified unique categories in the dataset: **Furniture**, **Office Supplies**, **Technology**.
2. Used `COUNTIF()` to count products per category:
   ```excel
   =COUNTIF(Orders!C:C, "Furniture")
   ```
3. Built a summary table (Category vs. Product Count).
4. Cross-verified counts using a **PivotTable** (Rows = Category, Values = Count of Category).
5. Used `INDEX` + `MATCH` to auto-identify the top category:
   ```excel
   =INDEX(A4:A6, MATCH(MAX(B4:B6), B4:B6, 0))
   ```

##  Result

| Category | Product Count |
|---|---|
| Furniture | 40 |
| Office Supplies | 38 |
| **Technology** | **42**  |

**Top Category:** Technology (42 products)

##  Deliverables

- [x] Product count table (category-wise)
- [x] Top category identified

##  Learning Outcome

Learned to summarize categorical data efficiently using `COUNTIF`/`COUNTIFS`, and to cross-verify formula-based results with a PivotTable.

---

**Author:** Akshat Srivastava
**Program:** Veda Technology — 45-Day Internship (Data Analytics Track)
