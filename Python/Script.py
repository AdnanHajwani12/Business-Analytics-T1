import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

orders = pd.read_excel("List_of_Orders_1.xlsx")
details = pd.read_excel("Order_Details_1.xlsx")
targets = pd.read_excel("Sales_target_1.xlsx")


orders.columns = orders.columns.str.strip()
details.columns = details.columns.str.strip()
targets.columns = targets.columns.str.strip()

orders.rename(columns={
    "Order ID": "Order_ID",
    "Order Date": "Order_Date",
    "CustomerName": "Customer_Name",
    "State": "State",
    "City": "City"
}, inplace=True)

details.rename(columns={
    "Order ID": "Order_ID",
    "Amount": "Amount",
    "Profit": "Profit",
    "Quantity": "Quantity",
    "Category": "Category",
    "Sub-Category": "Sub_Category"
}, inplace=True)

targets.rename(columns={
    "Month of Order Date": "Month",
    "Category": "Category",
    "Target": "Target"
}, inplace=True)


orders["Order_Date"] = pd.to_datetime(orders["Order_Date"], errors='coerce')
targets["Month"] = pd.to_datetime(targets["Month"], errors='coerce')
details["Amount"] = pd.to_numeric(details["Amount"], errors='coerce')
details["Profit"] = pd.to_numeric(details["Profit"], errors='coerce')


merged = pd.merge(details, orders, on="Order_ID", how="left")
category_summary = (
    merged.groupby("Category")
    .agg(
        Total_Sales=("Amount", "sum"),
        Total_Profit=("Profit", "sum"),
        Orders_Count=("Order_ID", "nunique")
    )
    .reset_index()
)

order_category = merged.groupby(["Order_ID", "Category"]).agg(
    Order_Sales=("Amount", "sum"),
    Order_Profit=("Profit", "sum")
).reset_index()

avg_profit = order_category.groupby("Category").agg(
    Avg_Profit_per_Order=("Order_Profit", "mean")
).reset_index()

category_summary = category_summary.merge(avg_profit, on="Category", how="left")
category_summary["Profit_Margin_%"] = (
    (category_summary["Total_Profit"] / category_summary["Total_Sales"]) * 100
).round(2)

top_performing = category_summary.sort_values("Profit_Margin_%", ascending=False).head(1)
under_performing = category_summary.sort_values("Profit_Margin_%", ascending=True).head(1)

furniture_targets = targets[targets["Category"].str.lower() == "furniture"].sort_values("Month")
furniture_targets["Target"] = pd.to_numeric(furniture_targets["Target"], errors='coerce')
furniture_targets["%_Change_MoM"] = furniture_targets["Target"].pct_change() * 100
furniture_targets["Significant_Change"] = furniture_targets["%_Change_MoM"].abs() > 20

top_states = (
    orders.groupby("State")
    .agg(Order_Count=("Order_ID", "nunique"))
    .sort_values("Order_Count", ascending=False)
    .head(5)
    .reset_index()
)
top5_list = top_states["State"].tolist()
regional_summary = (
    merged[merged["State"].isin(top5_list)]
    .groupby("State")
    .agg(
        Total_Sales=("Amount", "sum"),
        Avg_Profit=("Profit", "mean"),
        Order_Count=("Order_ID", "nunique")
    )
    .reset_index()
    .sort_values("Total_Sales", ascending=False)
)

output_dir = "Outputs"
os.makedirs(output_dir, exist_ok=True)
category_summary.to_excel(f"{output_dir}/Category_Sales_Profitability.xlsx", index=False)
furniture_targets.to_excel(f"{output_dir}/Furniture_Target_Analysis.xlsx", index=False)
regional_summary.to_excel(f"{output_dir}/Top5_Regional_Performance.xlsx", index=False)


plt.figure(figsize=(8, 5))
plt.bar(category_summary["Category"], category_summary["Total_Sales"])
plt.title("Total Sales by Category")
plt.xlabel("Category")
plt.ylabel("Sales (₹)")
plt.xticks(rotation=30)
plt.tight_layout()
plt.savefig(f"{output_dir}/Total_Sales_by_Category.png")
plt.close()

plt.figure(figsize=(8, 5))
plt.bar(category_summary["Category"], category_summary["Profit_Margin_%"], color="orange")
plt.title("Profit Margin % by Category")
plt.xlabel("Category")
plt.ylabel("Profit Margin (%)")
plt.xticks(rotation=30)
plt.tight_layout()
plt.savefig(f"{output_dir}/Profit_Margin_by_Category.png")
plt.close()
print("Analyzed!")
print(f"Results saved in folder: {output_dir}")
