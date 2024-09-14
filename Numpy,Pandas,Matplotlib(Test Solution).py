# Import necessary libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Task 1: Load the Data
online_retail = pd.read_excel(r"C:\Users\Lenovo\OneDrive\Pictures\Online Retail.xlsx")

# Task 2: Data Cleaning
print("Missing values:")
print(online_retail.isnull().sum())
online_retail.dropna(inplace=True)  # Remove rows with missing values
online_retail = online_retail[online_retail['Quantity'] > 0]  # Remove transactions with Quantity <= 0

# Task 3: Summary Statistics
online_retail['TotalRevenue'] = online_retail['Quantity'] * online_retail['UnitPrice']
product_summary = online_retail.groupby('StockCode')[['TotalRevenue', 'Quantity']].sum()  # Corrected column selection
print("Summary statistics:")
print(product_summary)

# Task 4: Monthly Revenue Analysis
online_retail['InvoiceDate'] = pd.to_datetime(online_retail['InvoiceDate'])
online_retail['Month'] = online_retail['InvoiceDate'].dt.strftime('%Y-%m')
monthly_revenue = online_retail.groupby('Month')['TotalRevenue'].sum()
print("Monthly revenue:")
print(monthly_revenue)
monthly_revenue.plot(kind='line', title='Monthly Revenue')
plt.show()

# Task 5: Top 5 Products
top_5_products = product_summary.nlargest(5, 'TotalRevenue')
print("Top 5 products:")
print(top_5_products)
top_5_products['TotalRevenue'].plot(kind='bar', title='Top 5 Products by Revenue')
plt.show()

# Task 6: Customer Analysis
customer_revenue = online_retail.groupby('CustomerID')['TotalRevenue'].sum()
print("Customer revenue:")
print(customer_revenue)
print("Customer with highest total revenue:")
print(customer_revenue.idxmax())

# Task 7: Quantity vs. Revenue Correlation
print("Correlation between Quantity and Total Revenue:")
print(np.corrcoef(online_retail['Quantity'], online_retail['TotalRevenue'])[0, 1])