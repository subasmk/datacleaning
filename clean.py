"""
Data Cleaning & Reporting Automation
Author: [subash m]
Date: June 2026
"""

import pandas as pd
from datetime import datetime

print("=" * 50)
print("DATA CLEANING AUTOMATION")
print("=" * 50)

# Load data
df = pd.read_csv('dirty_data.csv')
print(f"\n[1] Original data: {df.shape[0]} rows, {df.shape[1]} columns")

# Initial issues found
print(f"    - Duplicates: {df.duplicated().sum()}")
print(f"    - Missing values: {df.isnull().sum().sum()}")

# 1. Remove duplicates
df = df.drop_duplicates()
print(f"\n[2] After removing duplicates: {df.shape[0]} rows")

# 2. Handle missing values
df['sales'] = df['sales'].fillna(df['sales'].median())
df['name'] = df['name'].fillna('Unknown')
print(f"[3] Missing values filled (sales with median, name with 'Unknown')")

# 3. Standardize text
df['region'] = df['region'].str.lower().str.capitalize()
print(f"[4] Regions standardized: {df['region'].unique().tolist()}")

# 4. Fix date format
df['date'] = pd.to_datetime(df['date'], errors='coerce')
print(f"[5] Date format standardized")

# Save cleaned data
df.to_csv('cleaned_data.csv', index=False)
print(f"\n✅ SAVED: cleaned_data.csv")

# Generate Excel report
with pd.ExcelWriter('report.xlsx', engine='openpyxl') as writer:
    # Sheet 1: Cleaned data
    df.to_excel(writer, sheet_name='Cleaned Data', index=False)
    
    # Sheet 2: Summary statistics
    summary = pd.DataFrame({
        'Metric': ['Total Records', 'Total Columns', 'Date Range (Start)', 'Date Range (End)', 
                   'Average Sales', 'Median Sales', 'Total Sales', 'Unique Regions', 'Missing Values Fixed'],
        'Value': [
            len(df),
            len(df.columns),
            df['date'].min().strftime('%Y-%m-%d') if pd.notna(df['date'].min()) else 'N/A',
            df['date'].max().strftime('%Y-%m-%d') if pd.notna(df['date'].max()) else 'N/A',
            f"${df['sales'].mean():.2f}",
            f"${df['sales'].median():.2f}",
            f"${df['sales'].sum():.2f}",
            df['region'].nunique(),
            df.isnull().sum().sum()
        ]
    })
    summary.to_excel(writer, sheet_name='Summary', index=False)
    
    # Sheet 3: Sales by region
    region_stats = df.groupby('region').agg({
        'sales': ['count', 'mean', 'sum', 'min', 'max']
    }).round(2)
    region_stats.columns = ['Count', 'Avg Sales', 'Total Sales', 'Min', 'Max']
    region_stats.to_excel(writer, sheet_name='Sales by Region')
    
    # Sheet 4: Data quality report
    quality = pd.DataFrame({
        'Column': df.columns,
        'Data Type': df.dtypes.values,
        'Non-Null Count': df.count().values,
        'Null Count': df.isnull().sum().values,
        'Unique Values': df.nunique().values
    })
    quality.to_excel(writer, sheet_name='Data Quality', index=False)

print(f"✅ SAVED: report.xlsx")

# Final summary
print("\n" + "=" * 50)
print("FINAL SUMMARY")
print("=" * 50)
print(f"✓ Total clean records: {len(df)}")
print(f"✓ Average sales: ${df['sales'].mean():.2f}")
print(f"✓ Regions: {', '.join(df['region'].unique())}")
print(f"✓ Date range: {df['date'].min().date()} to {df['date'].max().date()}")
print("\n✅ Automation complete! Files ready for submission.")
