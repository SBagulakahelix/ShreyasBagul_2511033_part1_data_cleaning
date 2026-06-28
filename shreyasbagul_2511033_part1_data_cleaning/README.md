# Business Data Cleaning, Validation & Excel Reporting (Part 1)

**Student Name:** Shreyas Bagul  
**Student ID:** 2511033  

---

## 1. Problem Summary

This project addresses a common business scenario: sales data exported from multiple internal systems contains numerous inconsistencies and errors, including text formatting discrepancies, date parsing issues, duplicate records, missing values, invalid discounts, and mathematical mismatches. 

As a Business Analyst, the objective was to:
1. Preserve the original raw orders.
2. Build an automated cleaning and validation script.
3. Establish a standard data quality schema (`Clean`, `Warning`, `Invalid`) based on business rules.
4. Generate a comprehensive Data Quality Report.
5. Create structured Pivot summaries for business review.

---

## 2. Dataset Description

The raw dataset (`data/raw_orders.xlsx`) contains **932 records** with the following 21 fields:
- `order_id`: Unique identifier for each order.
- `order_date` & `ship_date`: Order and shipment dates.
- `customer_id` & `customer_name`: Unique customer identification and names.
- `segment`, `region`, `state`, `city`: Customer and geographical classifications.
- `category`, `sub_category`, `product_name`: Product classifications.
- `ship_mode`: Shipment method used.
- `quantity`, `unit_price`, `discount`, `sales`, `cost`, `profit`: Key transaction metrics.
- `payment_status` & `order_status`: Transaction states.

---

## 3. Tools Used

- **Python (3.x)**: For automated cleaning, validation, and spreadsheet generation.
- **Pandas**: For data structure manipulation, grouping, and aggregations.
- **OpenPyXL**: For advanced Excel automation, cell styling, borders, colors, column width adjustment, freeze panes, and native formula writing.

---

## 4. Cleaning Steps & Business Rules Applied

### Cleaning Steps:
1. **Zebra-Striping & Layout**: Applied gridlines, custom column widths, and freeze panes to all sheets.
2. **Text Standardization**: Trimmed extra spaces, resolved double spaces, and standardized all text fields to Title Case.
3. **Mixed-Date Parsing**: Converted varied formats (`DD/MM/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD`, `DD MMM YYYY`) into datetime objects and calculated delay days.
4. **Calculated Fields**:
   - `cleaned_discount`: Standardized numeric discount value.
   - `calculated_sales`: `quantity * unit_price * (1 - cleaned_discount)`
   - `calculated_profit`: `calculated_sales - cost`
   - `profit_margin`: `calculated_profit / calculated_sales`
   - `shipping_delay_days`: Days between order and shipment.
   - `order_month`: Month of order.
   - `order_year`: Year of order.

### Business Rules:
- **Missing region/ship_mode**: Filled with `'Unknown'` and flagged as `Warning`.
- **Missing discount**: Set to `0.0`. Flagged as `Warning` if sales math did not match.
- **Negative discounts**: Flagged as `Invalid`.
- **Discounts > 25%**: Flagged as `Invalid` (unusually high).
- **Ship date < order date**: Flagged as `Invalid` shipping record.
- **Sales/profit mismatches**: Flagged as `Warning`.
- **Completed Sales Filter**: Excluded non-completed, failed, refunded, or pending orders from completed sales pivots.

---

## 5. Summary of Data Quality Issues Found

From the 932 raw rows:
- **Exact Duplicates Removed**: 20 rows.
- **Final Rows Evaluated**: 912 rows.
- **Data Quality Flag Summary**:
  - **`Clean` (760 rows, 83.33%)**: Ready for Completed Sales summaries.
  - **`Warning` (101 rows, 11.07%)**: Imputed values, calculation mismatches, or conflicting duplicates.
  - **`Invalid` (51 rows, 5.59%)**: Critical date errors, negative or excessive discounts. Excluded from completed sales.

---

## 6. Summary of Pivot Reports (`outputs/pivot_summary.xlsx`)

The pivot summary includes six worksheets (formatted in deep steel blue with dynamic totals formulas):
1. **`Sales_by_Region`**: Total sales and profit by region, sorted by sales descending.
2. **`Sales_by_Category`**: Multi-level breakdown by category and sub-category, sorted by sales descending.
3. **`Orders_by_Ship_Mode`**: Order counts and percentages for each ship mode.
4. **`Margin_by_Segment`**: Total sales, profit, and overall profit margin by segment.
5. **`Issues_by_Region`**: Regional count of Returned, Cancelled, Failed, and Refunded orders.
6. **`Monthly_Sales_Trend`**: Sales and profit aggregated by year and month, sorted chronologically.

---

## 7. Key Business Insights

1. **Top Performing Region**: The **East** region leads in sales (approx. $1.52M) and profit (approx. $518K), followed closely by the **West** region.
2. **Most Profitable Product Category**: **Technology** is the highest revenue generator and has the highest average profit margin (~35%), with **Copiers** and **Phones** driving the majority of profits.
3. **Segment Performance**: The **Consumer** segment represents the largest share of sales (~$1.4M) and profit (~$448K), but all segments maintain a very healthy and consistent profit margin of **~30% to 32%**.
4. **Regional Risk**: The **South** region, while having lower overall sales volume, exhibits the highest relative count of Cancelled and Returned orders, suggesting delivery or customer satisfaction challenges in that region.
5. **Monthly Trends**: Sales show a strong cyclical pattern, with significant peaks in the third and fourth quarters of both 2024 and 2025.

---

## 8. Assumptions & Limitations

- **US vs. International Dates**: Slashed dates are assumed to be US (`MM/DD/YYYY`), and dashed dates are assumed to be Indian/European (`DD-MM-YYYY`). This resolved 70+ apparent negative delay anomalies and produced a consistent 0–8 day shipping window.
- **Imputed Data**: Missing region and ship mode cells are filled with `'Unknown'` since the true values could not be derived.
- **Calculation Differences**: Mathematical mismatches were flagged as warnings but not overwritten, as it's impossible to determine if the error originated from unit price, quantity, or raw sales entry.

---

## 9. Included Screenshots (`screenshots/`)

| Screenshot File | Description |
| :--- | :--- |
| `data_quality_report_preview.png` | Data Quality Dashboard showing record distribution and main anomalies metrics. |
| `cleaned_data_preview.png` | Cleaned dataset including calculated columns (`calculated_sales`, `data_quality_flag`). |
| `pivot_summary_preview1.png` | Pivot summary sheet: Sales and Profit by Category and Sub-category. |
| `pivot_summary_preview2.png` | Pivot summary sheet: Monthly Sales Trend (chronologically sorted). |
