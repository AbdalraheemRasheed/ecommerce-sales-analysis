import pandas as pd

df = pd.read_csv(r"C:\Users\hp\Desktop\Data analysis\Projects\sales_data_jan_jun_2024.csv")

df['Revenue'] = df['Quantity'] * df['Unit_Price']
df['Cost_Total'] = df['Quantity'] * df['Cost_Per_Unit']
df['Profit'] = df['Revenue'] - df['Cost_Total']

df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)

top_selling_products = df.groupby('Product')['Quantity'].sum().sort_values(ascending=False)
print("Top Selling Products:\n", top_selling_products)

monthly_sales = df.groupby(df['Date'].dt.to_period('M'))['Revenue'].sum()
print("\nMonthly Sales:\n", monthly_sales)

regional_performance = df.groupby('Region')['Revenue'].sum().sort_values(ascending=False)
print("\nRegional Performance:\n", regional_performance)

category_revenue = df.groupby('Category')['Revenue'].sum().sort_values(ascending=False)
print("\nRevenue by Category:\n", category_revenue)

total_revenue = category_revenue.sum()
category_revenue_pct = (category_revenue / total_revenue * 100).round(1)
print("\nRevenue Percentage by Category:\n", category_revenue_pct)

total_profit = df['Profit'].sum()
print(f"\nTotal Profit: ${total_profit:,.2f}")

category_profit = df.groupby('Category')['Profit'].sum().sort_values(ascending=False)
print("\nProfit by Category:\n", category_profit)

df.to_excel('sales_cleaned.xlsx', index=False)
print("\n✅ Data saved to 'sales_cleaned.xlsx'")