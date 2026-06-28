import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# Define file paths
project_dir = r"C:\Users\bagul\OneDrive\Desktop\assignment\shreyasbagul_2511033_part1_data_cleaning"
raw_path = os.path.join(project_dir, "data", "raw_orders.xlsx")
clean_path = os.path.join(project_dir, "data", "cleaned_orders.xlsx")
ss_dir = os.path.join(project_dir, "screenshots")

# Ensure screenshots directory exists
os.makedirs(ss_dir, exist_ok=True)

# Load datasets
df_raw = pd.read_excel(raw_path, sheet_name="raw_orders")
df_clean = pd.read_excel(clean_path, sheet_name="cleaned_orders")

# Configure Matplotlib styles
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'DejaVu Sans']

# Colors
header_color = '#1F4E79'   # Deep Steel Blue
zebra_color = '#F2F2F2'    # Light Gray
white_color = '#FFFFFF'
border_color = '#D9D9D9'
clean_color = '#E2EFDA'    # Soft Green
warning_color = '#FFF2CC'  # Soft Yellow
invalid_color = '#FCE4D6'  # Soft Red/Orange

def render_table(df_sample, title, filename, col_widths=None, highlight_dq_col=None, num_formats={}):
    fig, ax = plt.subplots(figsize=(14, 7), dpi=150)
    ax.axis('off')
    ax.axis('tight')
    
    # Format data for display
    display_data = []
    for row_idx, row in enumerate(df_sample.values):
        display_row = []
        for col_idx, val in enumerate(row):
            col_name = df_sample.columns[col_idx]
            # check formatting
            if col_name in num_formats:
                fmt = num_formats[col_name]
                if pd.isnull(val) or val == '-':
                    display_row.append('-')
                else:
                    display_row.append(fmt.format(val))
            else:
                if pd.isnull(val):
                    display_row.append('-')
                else:
                    display_row.append(str(val))
        display_data.append(display_row)
        
    table = ax.table(
        cellText=display_data, 
        colLabels=df_sample.columns, 
        loc='center', 
        cellLoc='center'
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1.0, 1.4)
    
    # Apply cell widths if specified
    if col_widths:
        for col_idx, width in col_widths.items():
            for row_idx in range(len(df_sample) + 1):
                table.get_celld()[row_idx, col_idx].set_width(width)
                
    # Style Cells
    for (row_idx, col_idx), cell in table.get_celld().items():
        cell.set_edgecolor(border_color)
        if row_idx == 0:
            # Header row
            cell.set_text_props(weight='bold', color=white_color)
            cell.set_facecolor(header_color)
        else:
            # Data rows
            cell.set_text_props(weight='normal')
            if row_idx % 2 == 0:
                cell.set_facecolor(zebra_color)
            else:
                cell.set_facecolor(white_color)
                
            # Highlight DQ Flag column
            if highlight_dq_col is not None and col_idx == highlight_dq_col:
                val = str(df_sample.iloc[row_idx - 1, col_idx])
                if val == 'Clean':
                    cell.set_facecolor(clean_color)
                elif val == 'Warning':
                    cell.set_facecolor(warning_color)
                elif val == 'Invalid':
                    cell.set_facecolor(invalid_color)
                cell.set_text_props(weight='bold')
                
    plt.title(title, fontsize=14, weight='bold', color='#1F4E79', pad=20)
    plt.tight_layout()
    
    out_path = os.path.join(ss_dir, filename)
    plt.savefig(out_path, bbox_inches='tight')
    plt.close()
    print(f"Table rendered and saved to {out_path}")

# ==========================================
# SCREENSHOT 1: data_quality_report_preview.png
# ==========================================
print("Rendering data_quality_report_preview.png...")
fig, ax = plt.subplots(figsize=(14, 8), dpi=150)
ax.axis('off')
ax.axis('tight')

# KPI Cards data
kpis = [
    ("Total Raw Records", "932"),
    ("Duplicates Removed", "20"),
    ("Clean (No Issues)", "760"),
    ("Warning (Minor Issues)", "101"),
    ("Invalid (Critical Issues)", "51")
]

# Draw KPI cards using text boxes
for idx, (label, val) in enumerate(kpis):
    x = 0.1 + idx * 0.20
    y = 0.82
    text_content = f"{label}\n\n{val}"
    ax.text(
        x, y, text_content, 
        ha='center', va='center', 
        fontsize=9, weight='bold', color='#1F4E79',
        bbox=dict(boxstyle="round,pad=1.2", facecolor='#DDEBF7', edgecolor='#1F4E79', lw=1.5)
    )

# Table data
df_dist = pd.DataFrame([
    ["Clean", "760", "83.33%", "No anomalies detected; ready for final completed sales summary."],
    ["Warning", "101", "11.07%", "Contains minor issues (missing region/ship_mode filled, sales mismatches, or conflicting duplicates). Included in analysis but marked for business review."],
    ["Invalid", "51", "5.59%", "Contains critical issues (negative/high discount, ship date before order date). Excluded from completed sales summaries."]
], columns=["Data Quality Flag", "Record Count", "Percentage", "Description / Action Taken"])

table = ax.table(
    cellText=df_dist.values,
    colLabels=df_dist.columns,
    loc='center',
    cellLoc='center',
    bbox=[0.02, 0.15, 0.96, 0.45]
)
table.auto_set_font_size(False)
table.set_fontsize(9)

# Style cells
for (row_idx, col_idx), cell in table.get_celld().items():
    cell.set_edgecolor('#D9D9D9')
    if row_idx == 0:
        # Header row
        cell.set_text_props(weight='bold', color='white')
        cell.set_facecolor('#1F4E79')
    else:
        # Data rows
        cell.set_text_props(weight='normal')
        if row_idx % 2 == 0:
            cell.set_facecolor('#F2F2F2')
        else:
            cell.set_facecolor('white')
            
        # Highlight DQ Flag column
        if col_idx == 0:
            val = df_dist.iloc[row_idx - 1, col_idx]
            if val == 'Clean':
                cell.set_facecolor('#E2EFDA')
            elif val == 'Warning':
                cell.set_facecolor('#FFF2CC')
            elif val == 'Invalid':
                cell.set_facecolor('#FCE4D6')
            cell.set_text_props(weight='bold')

plt.title("Data Quality Report Dashboard", fontsize=16, weight='bold', color='#1F4E79', pad=25)
plt.tight_layout()
plt.savefig(os.path.join(ss_dir, "data_quality_report_preview.png"), bbox_inches='tight')
plt.close()
print("Table rendered and saved to data_quality_report_preview.png")

# ==========================================
# SCREENSHOT 2: cleaned_data_preview.png
# ==========================================
print("Rendering cleaned_data_preview.png...")
# Select first 12 rows of cleaned orders, showing calculated columns
cols_clean_preview = ['order_id', 'order_date', 'customer_name', 'region', 'cleaned_discount', 'calculated_sales', 'calculated_profit', 'profit_margin', 'shipping_delay_days', 'data_quality_flag']
df_clean_preview = df_clean.head(12)[cols_clean_preview].copy()

num_formats_clean = {
    'cleaned_discount': '{:.1%}',
    'calculated_sales': '${:,.2f}',
    'calculated_profit': '${:,.2f}',
    'profit_margin': '{:.1%}'
}

render_table(
    df_clean_preview, 
    "Cleaned Dataset Preview (With Calculated Columns & DQ Flags)", 
    "cleaned_data_preview.png",
    highlight_dq_col=9, # data_quality_flag is index 9
    num_formats=num_formats_clean
)

# ==========================================
# SCREENSHOT 3: pivot_summary_preview1.png (Sales by Category)
# ==========================================
print("Rendering pivot_summary_preview1.png...")
# Aggregate Completed Sales by Category & Sub-category
df_completed = df_clean[(df_clean['data_quality_flag'] != 'Invalid') & (df_clean['order_status'] == 'Completed') & (df_clean['payment_status'] == 'Paid')].copy()
df_cat = df_completed.groupby(['category', 'sub_category'])[['calculated_sales', 'calculated_profit']].sum().reset_index()
df_cat = df_cat.sort_values(by=['category', 'calculated_sales'], ascending=[True, False]).head(12)

# Calculate profit margin
df_cat['profit_margin'] = df_cat['calculated_profit'] / df_cat['calculated_sales']

# Rename columns for presentation
df_cat.columns = ['Category', 'Sub-Category', 'Sales', 'Profit', 'Profit Margin']

num_formats_cat = {
    'Sales': '${:,.2f}',
    'Profit': '${:,.2f}',
    'Profit Margin': '{:.1%}'
}

render_table(
    df_cat, 
    "Pivot Table: Sales & Profit by Category & Sub-Category (Top 12)", 
    "pivot_summary_preview1.png",
    num_formats=num_formats_cat
)

# ==========================================
# SCREENSHOT 4: pivot_summary_preview2.png (Monthly Sales Trend)
# ==========================================
print("Rendering pivot_summary_preview2.png...")
# Group Completed Sales by Year & Month
month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
df_trend = df_completed.groupby(['order_year', 'order_month'])[['calculated_sales', 'calculated_profit']].sum().reset_index()
df_trend['month_num'] = df_trend['order_month'].apply(lambda m: month_order.index(m) if m in month_order else 99)
df_trend = df_trend.sort_values(by=['order_year', 'month_num']).drop(columns=['month_num']).head(12)

df_trend['profit_margin'] = df_trend['calculated_profit'] / df_trend['calculated_sales']
df_trend.columns = ['Year', 'Month', 'Sales', 'Profit', 'Profit Margin']

num_formats_trend = {
    'Sales': '${:,.2f}',
    'Profit': '${:,.2f}',
    'Profit Margin': '{:.1%}'
}

render_table(
    df_trend, 
    "Pivot Table: Monthly Sales & Profit Trend (First 12 Months)", 
    "pivot_summary_preview2.png",
    num_formats=num_formats_trend
)

print("Screenshots rendering completed successfully!")

