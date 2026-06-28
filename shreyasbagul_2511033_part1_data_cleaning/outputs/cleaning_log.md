# Data Cleaning Log - Part 1

**Student Name:** Shreyas Bagul  
**Student ID:** 2511033  
**Date:** June 22, 2026  

---

## 1. Summary of Dataset Issues Found

During the initial data profiling of the `raw_orders` sheet (932 rows, 21 columns), the following issues were identified:

| Issue Area | Description | Impact | Affected Rows |
| :--- | :--- | :--- | :---: |
| **Exact Duplicates** | Identical copies of the exact same row. | Skews sales & orders. | 20 |
| **Conflicting Duplicates** | Duplicate `order_id` values with conflicting `sales` and `order_status` values. | Inconsistent transaction record. | 12 groups (24 rows) |
| **Text Casing & Spaces** | Leading/trailing spaces, double spaces, and casing issues across all text columns. | Skews pivots (e.g. `'corporate'` vs `'Corporate'`). | Hundreds |
| **Date Parsing Issues** | Dates written in mixed formats (`DD/MM/YYYY`, `DD-MM-YYYY`, `DD MMM YYYY`, `YYYY-MM-DD`). | Python parses them incorrectly if default formats are used. | 700+ |
| **Date Sequences** | `ship_date` is earlier than `order_date` (negative shipping delay). | Invalid order lifecycle. | 21 |
| **Missing Values** | Missing entries in `region`, `ship_mode`, and `discount`. | Null values in report columns. | 64 |
| **Negative Discounts** | Recorded discounts below 0% (e.g. `-0.15`). | Negative sales deductions (invalid). | 15 |
| **Excessive Discounts** | Recorded discounts above allowed range (e.g. `55%`, `65%`, `70%`, `85%`). | Anomalously high discounts. | 15 |
| **Calculation Mismatches** | Raw `sales` and `profit` columns differ from recalculated values. | Skews financial results. | 64 |

---

## 2. Cleaning Actions & Logic Applied

### Step 1: Duplicate Handling
- **Action:** Removed 20 exact duplicate rows.
- **Logic:** Used `df.drop_duplicates()`, keeping the first occurrence.
- **Conflicting Duplicates:** Did **not** delete conflicting duplicate records. Instead, both rows for the 12 duplicate order IDs (24 rows) were preserved in `cleaned_orders.xlsx` and flagged as `Warning` for business review. A dedicated sheet was created in the Data Quality Report to list them.

### Step 2: Text Standardization
- **Action:** Cleaned the 10 text fields (`customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`).
- **Logic:**
  1. Applied `.strip()` to remove leading and trailing spaces.
  2. Substituted multiple spaces with a single space: `re.sub(r'\s+', ' ', val)`.
  3. Standardized capitalization to **Title Case** (e.g., `'office supplies'` -> `'Office Supplies'`, `'completed'` -> `'Completed'`).

### Step 3: Date Parsing & Standardization
- **Action:** Parsed and converted `order_date` and `ship_date` into datetime objects and output them in a consistent `YYYY-MM-DD` format.
- **Logic:**
  - Detected and parsed date strings conditionally:
    - `/` formatted dates parsed as `MM/DD/YYYY` (US format, e.g. `08/31/2024` -> `2024-08-31`)
    - `-` formatted dates parsed as `DD-MM-YYYY` (Indian/European format, e.g. `28-11-2024` -> `2024-11-28`)
    - `YYYY-MM-DD` parsed directly.
    - `DD MMM YYYY` (e.g. `21 Jul 2024`) parsed directly.
  - Calculated `shipping_delay_days = (ship_date - order_date).days`.
  - Flagged rows where `shipping_delay_days < 0` as `Invalid`.

### Step 4: Imputation of Missing Values
- **Region:** Filled 25 missing cells with `'Unknown'` and flagged as `Warning`.
- **Ship Mode:** Filled 21 missing cells with `'Unknown'` and flagged as `Warning`.
- **Discount:**
  - For the 18 missing values, compared raw `sales` with `quantity * unit_price`.
  - In 4 rows, the values matched, so the missing discount was treated as `0.0` and flagged as `Clean`.
  - In 14 rows, the values did not match (sales was lower, indicating a discount was applied but unrecorded). The discount was filled as `0.0` but flagged as `Warning` due to the math mismatch.

### Step 5: Discount Range Validation
- **Action:** Validated discounts.
- **Logic:**
  - Parsed string percentages (e.g., `'70%'` -> `0.70`).
  - Flagged negative discounts (`< 0.0`) as `Invalid`.
  - Flagged discounts above the allowed threshold (`> 0.25`) as `Invalid` (maximum valid retail discount in the cleaned dataset was 25%).

### Step 6: Recalculation & Margin Calculation
- **Calculated Columns:**
  - `cleaned_discount`: Standardized float discount.
  - `calculated_sales`: `quantity * unit_price * (1 - cleaned_discount)`
  - `calculated_profit`: `calculated_sales - cost`
  - `profit_margin`: `calculated_profit / calculated_sales`
  - `order_month`: Extracted month name (e.g., `July`).
  - `order_year`: Extracted year (e.g., `2024`).
- **Recalculation Flagging:** Flagged rows where `sales` or `profit` differed from recalculated values by more than `$0.05` as `Warning`.

---

## 3. Data Quality Flagging Rules

Every record in `cleaned_orders.xlsx` is tagged with a `data_quality_flag`:
1. **`Invalid` (Priority 1):** Critical data entry error. Excluded from completed sales summaries.
   - `shipping_delay_days < 0`
   - `cleaned_discount < 0`
   - `cleaned_discount > 0.25`
2. **`Warning` (Priority 2):** Minor issue or missing value that was imputed. Included in summaries but flagged for review.
   - `region == 'Unknown'`
   - `ship_mode == 'Unknown'`
   - Missing discount with sales calculation mismatch
   - Conflicting duplicate order ID
   - Sales/profit recalculation mismatch
3. **`Clean` (Priority 3):** No issues detected. Fully verified.

---

## 4. Key Metrics and Counts

- **Total Input Rows:** 932
- **Removed Rows (Exact Duplicates):** 20
- **Final Row Count in Cleaned File:** 912
- **Data Quality Flag Summary:**
  - **`Clean`:** 760 rows (83.33%)
  - **`Warning`:** 101 rows (11.07%)
  - **`Invalid`:** 51 rows (5.59%)
- **Completed Sales Filter:** Filtered to `order_status == 'Completed'` and `payment_status == 'Paid'` to get completed sales. This yields exactly **602 completed sales records**.

---

## 5. Assumptions and Limitations

### Assumptions:
1. **Mixed Date Formats:** Dates with slashes (`/`) were assumed to be in US format (`MM/DD/YYYY`), while dash-separated dates (`-`) were assumed to be in Indian/European format (`DD-MM-YYYY`). This was validated mathematically by examining the shipping delays, resulting in a clean 0–8 day shipping window.
2. **Imputation to 'Unknown':** Missing region and ship mode values are filled with `'Unknown'` since the correct values cannot be inferred from other fields.
3. **Completed Sales Definitive Filter:** Completed sales are defined as transactions that have `order_status == 'Completed'` and `payment_status == 'Paid'`. All other records (Returned, Cancelled, Failed, Refunded, Pending) are excluded from the main completed sales analysis pivots.

### Limitations:
1. **Mathematical Mismatches:** We recalculate sales and profit using unit price, quantity, and discount, but cannot determine if the mismatch in the raw data was due to incorrect raw sales entry, incorrect cost entry, or an unrecorded discount.
2. **Source System Errors:** Records flagged as `Invalid` cannot be automatically corrected and require manual investigation or database correction.
