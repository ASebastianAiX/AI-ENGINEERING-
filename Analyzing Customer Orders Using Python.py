"""
Course-end Project 1: Analyzing Customer Orders Using Python

What this script does:
1) Stores customer orders using Python data structures (lists/tuples/dicts/sets)
2) Computes:
   - Most frequently purchased products
   - Customer total spending + classification (High/Medium/Low)
   - Category profitability + top categories/products
   - Purchasing patterns (co-occurring products / “market basket” style)
3) Produces:
   - Clean tables (via pandas)
   - Charts (via matplotlib)
   - Excel report with multiple sheets
"""

from collections import defaultdict, Counter

# Optional reporting libraries (recommended)
import pandas as pd
import matplotlib.pyplot as plt


# -----------------------------------------------------------------------------
# 1) DATASET (Predefined Orders)
#    - You can replace or expand this data if your class provided another set.
# -----------------------------------------------------------------------------
# Each order is a dictionary:
# - order_id: string
# - customer_id: string
# - items: list of tuples (product_name, category, unit_price, quantity)
orders = [
    {
        "order_id": "O-1001",
        "customer_id": "C-001",
        "items": [
            ("Wireless Mouse", "Electronics", 25.00, 2),
            ("USB-C Cable", "Electronics", 10.00, 3),
            ("T-Shirt", "Clothing", 18.00, 1),
        ],
    },
    {
        "order_id": "O-1002",
        "customer_id": "C-002",
        "items": [
            ("Blender", "Home Essentials", 60.00, 1),
            ("Dish Soap", "Home Essentials", 5.00, 4),
        ],
    },
    {
        "order_id": "O-1003",
        "customer_id": "C-003",
        "items": [
            ("Jeans", "Clothing", 45.00, 1),
            ("Sneakers", "Clothing", 80.00, 1),
        ],
    },
    {
        "order_id": "O-1004",
        "customer_id": "C-001",
        "items": [
            ("Headphones", "Electronics", 120.00, 1),
            ("T-Shirt", "Clothing", 18.00, 2),
        ],
    },
    {
        "order_id": "O-1005",
        "customer_id": "C-004",
        "items": [
            ("Air Fryer", "Home Essentials", 95.00, 1),
            ("Wireless Mouse", "Electronics", 25.00, 1),
            ("Notebook", "Home Essentials", 6.00, 5),
        ],
    },
    {
        "order_id": "O-1006",
        "customer_id": "C-002",
        "items": [
            ("Headphones", "Electronics", 120.00, 1),
            ("USB-C Cable", "Electronics", 10.00, 2),
        ],
    },
]


# -----------------------------------------------------------------------------
# 2) HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def line_total(unit_price: float, qty: int) -> float:
    """Compute total for a single line item."""
    return unit_price * qty


def order_total(order: dict) -> float:
    """Compute total value of one order by summing its item totals."""
    total = 0.0
    for (product, category, unit_price, qty) in order["items"]:
        total += line_total(unit_price, qty)
    return total


def classify_customer(total_spend: float) -> str:
    """
    Classify customers based on total spending.
    Adjust thresholds to match your rubric if needed.
    """
    if total_spend >= 250:
        return "High-Value"
    elif total_spend >= 120:
        return "Mid-Value"
    else:
        return "Low-Value"


# -----------------------------------------------------------------------------
# 3) CORE ANALYSIS USING PYTHON DATA STRUCTURES
# -----------------------------------------------------------------------------

# A) Customer spending + order counts
customer_spend = defaultdict(float)     # customer_id -> total spend
customer_orders = defaultdict(int)      # customer_id -> number of orders

# B) Product frequency and revenue
product_qty = Counter()                # product -> total quantity sold
product_revenue = defaultdict(float)   # product -> total revenue

# C) Category revenue, units, customers
category_revenue = defaultdict(float)  # category -> total revenue
category_units = defaultdict(int)      # category -> total units sold
category_customers = defaultdict(set)  # category -> set of customers who bought from it

# D) “Market basket” / co-occurrence analysis using sets
# For each order, turn products into a set (unique products), then count pairs.
pair_counts = Counter()

# Process orders
order_totals = {}  # order_id -> total

for order in orders:
    cid = order["customer_id"]
    oid = order["order_id"]

    # Compute total per order (loops + conditionals)
    total = order_total(order)
    order_totals[oid] = total

    # Update customer aggregates
    customer_spend[cid] += total
    customer_orders[cid] += 1

    # Track set of products in this order for co-occurrence
    products_in_order = set()

    # Walk items
    for (product, category, unit_price, qty) in order["items"]:
        products_in_order.add(product)

        # product aggregates
        product_qty[product] += qty
        product_revenue[product] += line_total(unit_price, qty)

        # category aggregates
        category_revenue[category] += line_total(unit_price, qty)
        category_units[category] += qty
        category_customers[category].add(cid)

    # Count product pairs within the order
    # (simple pair counting; for small datasets this is fine)
    products_list = sorted(products_in_order)
    for i in range(len(products_list)):
        for j in range(i + 1, len(products_list)):
            pair = (products_list[i], products_list[j])
            pair_counts[pair] += 1


# -----------------------------------------------------------------------------
# 4) BUILD CUSTOMER CLASSIFICATION OUTPUT
# -----------------------------------------------------------------------------
customer_classification = []
for cid, spend in customer_spend.items():
    segment = classify_customer(spend)
    orders_count = customer_orders[cid]
    avg_order_value = spend / orders_count if orders_count else 0.0

    customer_classification.append({
        "customer_id": cid,
        "total_spend": round(spend, 2),
        "num_orders": orders_count,
        "avg_order_value": round(avg_order_value, 2),
        "segment": segment
    })

# Sort customers by spend (descending) to identify high-value
customer_classification.sort(key=lambda x: x["total_spend"], reverse=True)


# -----------------------------------------------------------------------------
# 5) PRODUCT + CATEGORY INSIGHTS
# -----------------------------------------------------------------------------
# Top products by quantity and by revenue
top_products_by_qty = product_qty.most_common(5)
top_products_by_revenue = sorted(product_revenue.items(), key=lambda x: x[1], reverse=True)[:5]

# Category profitability ranking
category_rank = sorted(category_revenue.items(), key=lambda x: x[1], reverse=True)


# -----------------------------------------------------------------------------
# 6) CO-OCCURRENCE (PURCHASING PATTERNS)
# -----------------------------------------------------------------------------
top_pairs = pair_counts.most_common(10)


# -----------------------------------------------------------------------------
# 7) CREATE REPORT TABLES (PANDAS) + CHARTS + EXCEL EXPORT
# -----------------------------------------------------------------------------
df_customers = pd.DataFrame(customer_classification)

df_products = pd.DataFrame([
    {"product": p, "units_sold": int(product_qty[p]), "revenue": round(product_revenue[p], 2)}
    for p in sorted(product_qty.keys())
]).sort_values(by="revenue", ascending=False)

df_categories = pd.DataFrame([
    {
        "category": c,
        "revenue": round(category_revenue[c], 2),
        "units_sold": int(category_units[c]),
        "unique_customers": len(category_customers[c]),
    }
    for c in category_revenue.keys()
]).sort_values(by="revenue", ascending=False)

df_pairs = pd.DataFrame([
    {"product_a": a, "product_b": b, "times_bought_together": int(cnt)}
    for (a, b), cnt in top_pairs
])

# ----- Print key outputs to console -----
print("\n=== CUSTOMER CLASSIFICATION (sorted by total_spend) ===")
print(df_customers.to_string(index=False))

print("\n=== TOP PRODUCTS BY QUANTITY ===")
print(pd.DataFrame(top_products_by_qty, columns=["product", "units_sold"]).to_string(index=False))

print("\n=== TOP PRODUCTS BY REVENUE ===")
print(pd.DataFrame(top_products_by_revenue, columns=["product", "revenue"]).to_string(index=False))

print("\n=== CATEGORY PROFITABILITY (by revenue) ===")
print(df_categories.to_string(index=False))

print("\n=== TOP CO-OCCURRING PRODUCT PAIRS ===")
print(df_pairs.to_string(index=False))


# -----------------------------------------------------------------------------
# 8) CHARTS (matplotlib)
# -----------------------------------------------------------------------------
# Category revenue bar chart
plt.figure()
plt.bar(df_categories["category"], df_categories["revenue"])
plt.title("Revenue by Category")
plt.xlabel("Category")
plt.ylabel("Revenue")
plt.xticks(rotation=30, ha="right")
plt.tight_layout()
plt.show()

# Top 5 products by revenue chart
top5 = df_products.head(5)
plt.figure()
plt.bar(top5["product"], top5["revenue"])
plt.title("Top 5 Products by Revenue")
plt.xlabel("Product")
plt.ylabel("Revenue")
plt.xticks(rotation=30, ha="right")
plt.tight_layout()
plt.show()

# Customer total spend chart
plt.figure()
plt.bar(df_customers["customer_id"], df_customers["total_spend"])
plt.title("Total Spend by Customer")
plt.xlabel("Customer")
plt.ylabel("Total Spend")
plt.tight_layout()
plt.show()


# -----------------------------------------------------------------------------
# 9) EXPORT AN EXCEL REPORT (Spreadsheet)
# -----------------------------------------------------------------------------
output_file = "customer_orders_report.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_customers.to_excel(writer, sheet_name="Customers", index=False)
    df_products.to_excel(writer, sheet_name="Products", index=False)
    df_categories.to_excel(writer, sheet_name="Categories", index=False)
    df_pairs.to_excel(writer, sheet_name="Product_Pairs", index=False)

print(f"\nExcel report saved to: {output_file}")


# -----------------------------------------------------------------------------
# 10) FINAL “BUSINESS INSIGHTS” (write these into your submission report)
# -----------------------------------------------------------------------------
# Here we create a concise, manager-friendly insight section.
insights = []

# 1) High-value customer identification
if not df_customers.empty:
    top_customer = df_customers.iloc[0]
    insights.append(
        f"Top customer is {top_customer['customer_id']} with total spend ${top_customer['total_spend']:.2f} "
        f"across {int(top_customer['num_orders'])} orders (segment: {top_customer['segment']})."
    )

# 2) Most profitable category
if not df_categories.empty:
    top_cat = df_categories.iloc[0]
    insights.append(
        f"Most profitable category is {top_cat['category']} with revenue ${top_cat['revenue']:.2f}. "
        f"It sold {int(top_cat['units_sold'])} units to {int(top_cat['unique_customers'])} unique customers."
    )

# 3) Most frequent product & highest revenue product
if top_products_by_qty:
    insights.append(f"Most frequently purchased product (by units) is '{top_products_by_qty[0][0]}' "
                    f"with {top_products_by_qty[0][1]} units sold.")
if top_products_by_revenue:
    insights.append(f"Highest revenue product is '{top_products_by_revenue[0][0]}' "
                    f"with ${top_products_by_revenue[0][1]:.2f} revenue.")

# 4) Cross-sell opportunity from co-occurrence
if top_pairs:
    a, b = top_pairs[0][0]
    cnt = top_pairs[0][1]
    insights.append(
        f"Top cross-sell pair is '{a}' + '{b}' (bought together {cnt} times). "
        f"Consider bundling or recommending these together at checkout."
    )

print("\n=== BUSINESS INSIGHTS SUMMARY ===")
for i, s in enumerate(insights, start=1):
    print(f"{i}. {s}")
