import pandas as pd
import numpy as np
import re
import os

# Define file paths
project_dir = r"C:\Users\bagul\OneDrive\Desktop\assignment\shreyasbagul_2511033_part1_data_cleaning"
raw_path = os.path.join(project_dir, "data", "raw_orders.xlsx")
clean_path = os.path.join(project_dir, "data", "cleaned_orders.xlsx")

print("--- Loading Raw Orders ---")
df = pd.read_excel(raw_path, sheet_name="raw_orders")
print(f"Raw data loaded. Shape: {df.shape}")

# Keep track of original values for reporting
df_orig = df.copy()

# ==========================================
# 1. Handle Exact Duplicates
# ==========================================
exact_duplicates_mask = df.duplicated()
exact_duplicate_count = exact_duplicates_mask.sum()
print(f"Exact duplicate rows found: {exact_duplicate_count}")

# Drop exact duplicates
df_clean = df.drop_duplicates().copy()
print(f"Shape after dropping exact duplicates: {df_clean.shape}")

# ==========================================
# 2. Text Cleaning & Standardization
# ==========================================
text_cols = ['customer_name', 'segment', 'region', 'state', 'city', 'category', 'sub_category', 'ship_mode', 'payment_status', 'order_status']

def clean_text_field(val):
    if pd.isnull(val):
        return val
    # Remove leading/trailing whitespace and replace multiple spaces with single space
    cleaned = str(val).strip()
    cleaned = re.sub(r'\s+', ' ', cleaned)
    # Standardize to Title Case (excluding nan)
    return cleaned.title()

for col in text_cols:
    df_clean[col] = df_clean[col].apply(clean_text_field)

print("Text columns cleaned and standardized.")

# ==========================================
# 3. Date Handling
# ==========================================
def parse_custom_date(val):
    if pd.isnull(val):
        return pd.NaT
    val = str(val).strip()
    
    # Format: YYYY-MM-DD
    if re.match(r'^\d{4}-\d{2}-\d{2}$', val):
        return pd.to_datetime(val, format='%Y-%m-%d', errors='coerce')
    # Format: DD-MM-YYYY (European/Indian style)
    elif re.match(r'^\d{2}-\d{2}-\d{4}$', val):
        return pd.to_datetime(val, format='%d-%m-%Y', errors='coerce')
    # Format: MM/DD/YYYY (US style)
    elif re.match(r'^\d{2}/\d{2}/\d{4}$', val):
        return pd.to_datetime(val, format='%m/%d/%Y', errors='coerce')
    # Format: DD MMM YYYY (e.g. 21 Jul 2024)
    elif re.match(r'^\d{1,2} [A-Za-z]{3} \d{4}$', val):
        return pd.to_datetime(val, format='%d %b %Y', errors='coerce')
    # Format: MMM DD, YYYY
    elif re.match(r'^[A-Za-z]{3} \d{1,2}, \d{4}$', val):
        return pd.to_datetime(val, format='%b %d, %Y', errors='coerce')
    else:
        return pd.to_datetime(val, errors='coerce')

# Parse and convert to datetime objects
df_clean['order_date_parsed'] = df_clean['order_date'].apply(parse_custom_date)
df_clean['ship_date_parsed'] = df_clean['ship_date'].apply(parse_custom_date)

# Calculate shipping delay in days
df_clean['shipping_delay_days'] = (df_clean['ship_date_parsed'] - df_clean['order_date_parsed']).dt.days

# Check for date issues
date_issues_mask = df_clean['shipping_delay_days'] < 0
print(f"Records with ship date before order date: {date_issues_mask.sum()}")

# Overwrite original date columns with cleaned date strings (YYYY-MM-DD)
df_clean['order_date'] = df_clean['order_date_parsed'].dt.strftime('%Y-%m-%d')
df_clean['ship_date'] = df_clean['ship_date_parsed'].dt.strftime('%Y-%m-%d')

# Add month and year calculated columns
df_clean['order_month'] = df_clean['order_date_parsed'].dt.strftime('%B')
df_clean['order_year'] = df_clean['order_date_parsed'].dt.year

# ==========================================
# 4. Handle Missing Values
# ==========================================
# Missing region -> Unknown
missing_region_mask = df_clean['region'].isnull()
df_clean.loc[missing_region_mask, 'region'] = 'Unknown'
print(f"Missing region rows filled with 'Unknown': {missing_region_mask.sum()}")

# Missing ship_mode -> Unknown
missing_ship_mode_mask = df_clean['ship_mode'].isnull()
df_clean.loc[missing_ship_mode_mask, 'ship_mode'] = 'Unknown'
print(f"Missing ship_mode rows filled with 'Unknown': {missing_ship_mode_mask.sum()}")

# ==========================================
# 5. Discount Handling
# ==========================================
def parse_discount(val):
    if pd.isnull(val):
        return None
    if isinstance(val, str):
        val = val.strip()
        if val.endswith('%'):
            return float(val.replace('%', '')) / 100.0
        try:
            return float(val)
        except ValueError:
            return None
    return float(val)

# Raw parsing of existing discounts
df_clean['raw_parsed_discount'] = df_clean['discount'].apply(parse_discount)

# Apply missing discount business rules:
# Treat as 0 only if all other sales fields are valid (sales == quantity * unit_price); otherwise flag.
missing_discount_mask = df_clean['discount'].isnull()
sales_fields_valid = (df_clean['sales'] - (df_clean['quantity'] * df_clean['unit_price'])).abs() <= 0.05

# In df_clean, we set cleaned_discount
df_clean['cleaned_discount'] = df_clean['raw_parsed_discount']

# For missing discounts:
# If other sales fields are valid, set discount to 0
# If other sales fields are NOT valid, set discount to 0 but we will flag it in the next section
df_clean.loc[missing_discount_mask, 'cleaned_discount'] = 0.0

# Check for discount issues
negative_discount_mask = df_clean['cleaned_discount'] < 0
high_discount_mask = df_clean['cleaned_discount'] > 0.25
missing_discount_invalid_mask = missing_discount_mask & (~sales_fields_valid)

print(f"Missing discount rows: {missing_discount_mask.sum()}")
print(f"  Of which, math matches 0% discount: {(missing_discount_mask & sales_fields_valid).sum()}")
print(f"  Of which, math does NOT match 0% discount (flagged): {missing_discount_invalid_mask.sum()}")
print(f"Negative discounts (< 0): {negative_discount_mask.sum()}")
print(f"High discounts (> 0.25): {high_discount_mask.sum()}")

# ==========================================
# 6. Duplicate Order ID Handling
# ==========================================
# Identify duplicate order IDs
dup_order_id_series = df_clean[df_clean.duplicated(subset=['order_id'], keep=False)]['order_id'].unique()
is_conflicting_duplicate = df_clean['order_id'].isin(dup_order_id_series)
print(f"Conflicting duplicate order IDs found: {len(dup_order_id_series)} (affecting {is_conflicting_duplicate.sum()} rows)")

# ==========================================
# 7. Create Calculated Columns
# ==========================================
# calculated_sales = quantity * unit_price * (1 - cleaned_discount)
df_clean['calculated_sales'] = df_clean['quantity'] * df_clean['unit_price'] * (1.0 - df_clean['cleaned_discount'])
# calculated_profit = calculated_sales - cost
df_clean['calculated_profit'] = df_clean['calculated_sales'] - df_clean['cost']
# profit_margin = profit / sales
df_clean['profit_margin'] = df_clean['calculated_profit'] / df_clean['calculated_sales']

# Check for calculation mismatches (diff > 0.05) between raw and calculated
sales_mismatch_mask = (df_clean['sales'] - df_clean['calculated_sales']).abs() > 0.05
profit_mismatch_mask = (df_clean['profit'] - df_clean['calculated_profit']).abs() > 0.05
calculation_mismatch = sales_mismatch_mask | profit_mismatch_mask

print(f"Calculation mismatches found: Sales={sales_mismatch_mask.sum()}, Profit={profit_mismatch_mask.sum()}")

# ==========================================
# 8. Data Quality Flag
# ==========================================
# Initialize as 'Clean'
df_clean['data_quality_flag'] = 'Clean'

# Assign 'Warning' conditions:
warning_mask = (
    missing_region_mask | 
    missing_ship_mode_mask | 
    missing_discount_invalid_mask | 
    is_conflicting_duplicate |
    calculation_mismatch
)
df_clean.loc[warning_mask, 'data_quality_flag'] = 'Warning'

# Assign 'Invalid' conditions (takes precedence over Warning):
invalid_mask = (
    date_issues_mask | 
    negative_discount_mask | 
    high_discount_mask
)
df_clean.loc[invalid_mask, 'data_quality_flag'] = 'Invalid'

# Summarize data quality counts
print("\nData Quality Flag Counts:")
print(df_clean['data_quality_flag'].value_counts())

# ==========================================
# 9. Clean up DataFrame Columns for Output
# ==========================================
# Let's see the columns we need to write:
# order_id, order_date, ship_date, customer_id, customer_name, segment, region, state, city,
# category, sub_category, product_name, ship_mode, quantity, unit_price, discount, sales, cost, profit,
# payment_status, order_status, cleaned_discount, calculated_sales, calculated_profit, profit_margin,
# shipping_delay_days, order_month, order_year, data_quality_flag

# Drop the temp columns we created
output_cols = [
    'order_id', 'order_date', 'ship_date', 'customer_id', 'customer_name', 'segment', 'region', 'state', 'city',
    'category', 'sub_category', 'product_name', 'ship_mode', 'quantity', 'unit_price', 'discount', 'sales', 'cost', 'profit',
    'payment_status', 'order_status', 'cleaned_discount', 'calculated_sales', 'calculated_profit', 'profit_margin',
    'shipping_delay_days', 'order_month', 'order_year', 'data_quality_flag'
]

# Ensure they exist
for col in output_cols:
    if col not in df_clean.columns:
        print(f"Warning: {col} not found in df_clean!")

df_output = df_clean[output_cols]

# Write to excel
print(f"\nSaving cleaned dataset to {clean_path}...")
# We will save the cleaned_orders with two sheets: 'cleaned_orders' and 'business_rules' (to preserve the structure)
with pd.ExcelWriter(clean_path, engine='openpyxl') as writer:
    df_output.to_excel(writer, sheet_name='cleaned_orders', index=False)
    # copy the business_rules sheet from original workbook
    xl_orig = pd.ExcelFile(raw_path)
    if 'business_rules' in xl_orig.sheet_names:
        df_rules = xl_orig.parse('business_rules')
        df_rules.to_excel(writer, sheet_name='business_rules', index=False)

print("Cleaned dataset saved successfully!")
