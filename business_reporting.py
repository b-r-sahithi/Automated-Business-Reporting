import pandas as pd
import numpy as np

# Load data
df = pd.read_csv("sales_data.csv")

# Basic exploration
print(df.info())
print(df.isnull().sum())

# Data cleaning
df = df.drop_duplicates()
df['revenue'] = df['revenue'].fillna(df['revenue'].median())

# Data validation
df = df[df['revenue'] > 0]
df = df[df['quantity'] > 0]

# Feature engineering
df['order_date'] = pd.to_datetime(df['order_date'])
df['month'] = df['order_date'].dt.to_period('M')

# KPI calculations
summary = {
    "Total Revenue": df['revenue'].sum(),
    "Total Orders": df['order_id'].nunique(),
    "Average Order Value": df['revenue'].mean()
}

kpi_df = pd.DataFrame(list(summary.items()), columns=['Metric', 'Value'])

# Aggregations for reporting
monthly_revenue = df.groupby('month')['revenue'].sum().reset_index()
category_revenue = df.groupby('category')['revenue'].sum().reset_index()

# Export to Excel
with pd.ExcelWriter("Business_Report.xlsx", engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Clean_Data', index=False)
    kpi_df.to_excel(writer, sheet_name='KPI_Summary', index=False)
    monthly_revenue.to_excel(writer, sheet_name='Monthly_Revenue', index=False)
    category_revenue.to_excel(writer, sheet_name='Category_Revenue', index=False)

print("Report generated successfully.")
