
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# =====================================================
# COFFEE SHOP BUSINESS ANALYTICS
# Data Cleaning, Exploration, and Business Insights
# =====================================================

BASE_DIR = Path(__file__).resolve().parent
FILE_PATH = BASE_DIR / "coffee_shop_data.xlsx"

# -----------------------------
# 1. LOAD EXCEL DATA
# -----------------------------

if not FILE_PATH.exists():
    raise FileNotFoundError(
        f"Excel file not found: {FILE_PATH}\n"
        "Place coffee_shop_data.xlsx in the same folder as this script."
    )

excel_file = pd.ExcelFile(FILE_PATH)

print("=" * 60)
print("COFFEE SHOP BUSINESS ANALYTICS")
print("=" * 60)
print("\nAvailable worksheets:", excel_file.sheet_names)

# Read sheets. The first row in Customers and Inventory
# contains their column names, despite Excel importing
# them as ordinary data.
sales = pd.read_excel(excel_file, sheet_name="sales", header=0)
customers = pd.read_excel(excel_file, sheet_name="customers", header=0)
inventory = pd.read_excel(excel_file, sheet_name="inventory", header=0)


# -----------------------------
# 2. FIX CUSTOMERS / INVENTORY HEADERS
# -----------------------------

def fix_header_row(df):
    if "Column1" in df.columns:
        df.columns = df.iloc[0].astype(str).str.strip()
        df = df.iloc[1:].reset_index(drop=True)
    return df


customers = fix_header_row(customers)
inventory = fix_header_row(inventory)


# -----------------------------
# 3. CLEAN COLUMN NAMES / TEXT
# -----------------------------

for df in [sales, customers, inventory]:
    df.columns = df.columns.astype(str).str.strip()

    # Trim whitespace from string values only.
    # This safely leaves dates and time values unchanged.
    for col in df.columns:
        df[col] = df[col].map(
            lambda value: value.strip()
            if isinstance(value, str)
            else value
        )

# Standardize a known product spelling variation.
sales["product"] = sales["product"].replace({
    "Cappucino": "Cappuccino"
})

# -----------------------------
# 4. CONVERT DATA TYPES
# -----------------------------

sales["order_date"] = pd.to_datetime(
    sales["order_date"], errors="coerce"
)

customers["signup_date"] = pd.to_datetime(
    customers["signup_date"], errors="coerce"
)

# Convert order times into a consistent time format.
sales["order_time"] = pd.to_datetime(
    sales["order_time"].astype(str),
    format="mixed",
    errors="coerce"
).dt.time

sales_numeric = [
    "quantity", "unit_price", "cost_per_unit",
    "Revenue", "Gross_Profit"
]

for col in sales_numeric:
    if col in sales.columns:
        sales[col] = pd.to_numeric(sales[col], errors="coerce")

inventory_numeric = [
    "stock_on_hand", "reorder_level",
    "monthly_units_sold", "units_wasted", "cost_per_unit"
]

for col in inventory_numeric:
    if col in inventory.columns:
        inventory[col] = pd.to_numeric(inventory[col], errors="coerce")

customers["age"] = pd.to_numeric(
    customers["age"], errors="coerce"
)

# -----------------------------
# 5. DATA QUALITY CHECKS
# -----------------------------

print("\n" + "=" * 60)
print("1. DATA QUALITY CHECKS")
print("=" * 60)

for name, df in [
    ("Sales", sales),
    ("Customers", customers),
    ("Inventory", inventory)
]:
    print(f"\n{name}: {len(df)} rows")
    print("Missing values by column:")
    print(df.isnull().sum().to_string())
    print("Exact duplicate rows:", df.duplicated().sum())

# Remove exact duplicate rows only.
sales = sales.drop_duplicates().copy()
customers = customers.drop_duplicates().copy()
inventory = inventory.drop_duplicates().copy()

# Recalculate financial metrics from the source columns.
sales["Revenue"] = sales["quantity"] * sales["unit_price"]

sales["Gross_Profit"] = (
    sales["quantity"]
    * (sales["unit_price"] - sales["cost_per_unit"])
)

print("\nRows after removing exact duplicates:")
print("Sales:", len(sales))
print("Customers:", len(customers))
print("Inventory:", len(inventory))


# -----------------------------
# 6. OVERALL SALES PERFORMANCE
# -----------------------------

print("\n" + "=" * 60)
print("2. OVERALL SALES PERFORMANCE")
print("=" * 60)

total_revenue = sales["Revenue"].sum()
total_gross_profit = sales["Gross_Profit"].sum()
total_units = sales["quantity"].sum()
average_transaction_value = sales["Revenue"].mean()

print(f"Total revenue: {total_revenue:,.2f}")
print(f"Total gross profit: {total_gross_profit:,.2f}")
print(f"Total units sold: {total_units}")
print(f"Average transaction value: {average_transaction_value:,.2f}")

if total_revenue:
    gross_margin = total_gross_profit / total_revenue * 100
    print(f"Gross profit margin: {gross_margin:.2f}%")


# -----------------------------
# 7. PRODUCT PERFORMANCE
# -----------------------------

print("\n" + "=" * 60)
print("3. PRODUCT PERFORMANCE")
print("=" * 60)

product_summary = sales.groupby("product").agg(
    units_sold=("quantity", "sum"),
    revenue=("Revenue", "sum"),
    gross_profit=("Gross_Profit", "sum")
).sort_values("revenue", ascending=False)

print("\nProducts ranked by revenue:")
print(product_summary.to_string())

if not product_summary.empty:
    print("\nHighest-revenue product:", product_summary["revenue"].idxmax())
    print(
        "Highest-gross-profit product:",
        product_summary["gross_profit"].idxmax()
    )

    product_summary["revenue"].sort_values().plot(
        kind="barh",
        title="Revenue by Product",
        figsize=(9, 6)
    )
    plt.xlabel("Revenue")
    plt.tight_layout()
    plt.savefig(BASE_DIR / "revenue_by_product.png", dpi=150)
    plt.show()


# -----------------------------
# 8. SALES TRENDS
# -----------------------------

print("\n" + "=" * 60)
print("4. SALES TRENDS")
print("=" * 60)

sales["day_name"] = sales["order_date"].dt.day_name()

sales["hour"] = sales["order_time"].map(
    lambda value: value.hour if pd.notna(value) else pd.NA
)

daily_sales = sales.groupby("order_date")["Revenue"].sum()
weekday_sales = sales.groupby("day_name")["Revenue"].sum()
hourly_sales = sales.dropna(subset=["hour"]).groupby("hour")["Revenue"].sum()

print("\nRevenue by date:")
print(daily_sales.to_string())

print("\nRevenue by weekday:")
print(weekday_sales.to_string())

print("\nRevenue by hour:")
print(hourly_sales.to_string())

if not daily_sales.empty:
    print("\nHighest-revenue date:", daily_sales.idxmax().date())

if not weekday_sales.empty:
    print("Highest-revenue weekday:", weekday_sales.idxmax())

if not hourly_sales.empty:
    print("Highest-revenue hour:", int(hourly_sales.idxmax()))

if not daily_sales.empty:
    daily_sales.plot(
        kind="line",
        marker="o",
        title="Daily Revenue Trend",
        figsize=(9, 5)
    )
    plt.ylabel("Revenue")
    plt.tight_layout()
    plt.savefig(BASE_DIR / "daily_sales_trend.png", dpi=150)
    plt.show()


# -----------------------------
# 9. CUSTOMER BEHAVIOUR
# -----------------------------

print("\n" + "=" * 60)
print("5. CUSTOMER BEHAVIOUR")
print("=" * 60)

customer_summary = sales.groupby("customer_id").agg(
    transactions=("order_id", "nunique"),
    units_purchased=("quantity", "sum"),
    total_spent=("Revenue", "sum"),
    total_gross_profit=("Gross_Profit", "sum"),
    last_purchase=("order_date", "max")
).sort_values("total_spent", ascending=False)

print("\nCustomer purchase summary:")
print(customer_summary.to_string())

if not customer_summary.empty:
    print(
        "\nHighest-spending customer ID:",
        customer_summary["total_spent"].idxmax()
    )

# Attach customer names and loyalty tiers.
customer_details = customers.set_index("customer_id")

customer_report = customer_summary.join(
    customer_details[["customer_name", "loyalty_tier"]],
    how="left"
)

print("\nCustomer summary with names and loyalty tiers:")
print(customer_report.to_string())

# Note: A short sample cannot reliably prove customer churn.
# Last purchase date only indicates recency within this dataset.


# -----------------------------
# 10. INVENTORY ANALYSIS
# -----------------------------

print("\n" + "=" * 60)
print("6. INVENTORY ANALYSIS")
print("=" * 60)

inventory["stock_status"] = inventory.apply(
    lambda row: (
        "Understocked"
        if row["stock_on_hand"] < row["reorder_level"]
        else "At or above reorder level"
    ),
    axis=1
)

# Waste rate = wasted units / (units sold + wasted units).
total_units_handled = (
    inventory["monthly_units_sold"] + inventory["units_wasted"]
)

inventory["waste_rate_percent"] = (
    inventory["units_wasted"]
    .div(total_units_handled.where(total_units_handled != 0))
    * 100
)

inventory_report_columns = [
    "product_name",
    "stock_on_hand",
    "reorder_level",
    "monthly_units_sold",
    "units_wasted",
    "waste_rate_percent",
    "stock_status"
]

print("\nInventory report:")
print(
    inventory[inventory_report_columns].to_string(index=False)
)

understocked = inventory[
    inventory["stock_on_hand"] < inventory["reorder_level"]
]

print("\nProducts below reorder level:")
if understocked.empty:
    print("None")
else:
    print(
        understocked[
            ["product_name", "stock_on_hand", "reorder_level"]
        ].to_string(index=False)
    )

print("\nProducts ranked by units wasted:")
print(
    inventory.sort_values("units_wasted", ascending=False)[
        ["product_name", "units_wasted", "waste_rate_percent"]
    ].to_string(index=False)
)


# -----------------------------
# 11. SAVE RESULTS TO EXCEL
# -----------------------------

output_file = BASE_DIR / "coffee_shop_analysis_results.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    product_summary.to_excel(writer, sheet_name="Product Performance")
    daily_sales.to_frame("Revenue").to_excel(writer, sheet_name="Daily Sales")
    weekday_sales.to_frame("Revenue").to_excel(
        writer, sheet_name="Weekday Sales"
    )
    hourly_sales.to_frame("Revenue").to_excel(
        writer, sheet_name="Hourly Sales"
    )
    customer_report.to_excel(writer, sheet_name="Customer Analysis")
    inventory.to_excel(
        writer, sheet_name="Inventory Analysis", index=False
    )

print("\n" + "=" * 60)
print("ANALYSIS COMPLETE")
print("=" * 60)
print(f"Excel results saved to:\n{output_file}")
print("Charts saved in the project folder.")
