import pandas as pd


def generate_analysis(df_sales):

    # =========================
    # TOTAL REVENUE
    # =========================
    total_revenue = df_sales["Revenue"].sum()

    # =========================
    # OPTIONAL COST / PROFIT METRICS
    # Only valid where cost exists
    # =========================
    valid_cost_rows = df_sales[df_sales["Cost"].notna()].copy()

    if not valid_cost_rows.empty:
        total_cost = valid_cost_rows["Cost"].sum()
        total_profit = valid_cost_rows["Profit"].sum()
        avg_margin = valid_cost_rows["Margin"].mean()
    else:
        total_cost = None
        total_profit = None
        avg_margin = None

    # =========================
    # REVENUE BY PRODUCT
    # =========================
    revenue_by_product = (
        df_sales.groupby("Product_Name")["Revenue"]
        .sum()
        .reset_index()
        .sort_values("Revenue", ascending=False)
    )

    # =========================
    # REVENUE BY REGION
    # =========================
    revenue_by_region = (
        df_sales.groupby("Region")["Revenue"]
        .sum()
        .reset_index()
        .sort_values("Revenue", ascending=False)
    )

    # =========================
    # MONTHLY TREND (FIXED ORDER)
    # =========================
    monthly_trend = (
        df_sales.groupby(["Month_Num", "Month"])["Revenue"]
        .sum()
        .reset_index()
        .sort_values("Month_Num")
    )

    monthly_trend = monthly_trend.drop(columns=["Month_Num"])

    # =========================
    # TOP CUSTOMERS
    # =========================
    top_customers = (
        df_sales.groupby("Customer_Name")["Revenue"]
        .sum()
        .reset_index()
        .sort_values("Revenue", ascending=False)
        .head(5)
    )

    # =========================
    # PROFIT BY PRODUCT
    # Only for products with known cost
    # =========================
    if not valid_cost_rows.empty:
        profit_by_product = (
            valid_cost_rows.groupby("Product_Name")["Profit"]
            .sum()
            .reset_index()
            .sort_values("Profit", ascending=False)
        )
    else:
        profit_by_product = pd.DataFrame(columns=["Product_Name", "Profit"])

    return {
        "total_revenue": total_revenue,
        "total_cost": total_cost,
        "total_profit": total_profit,
        "avg_margin": avg_margin,
        "revenue_by_product": revenue_by_product,
        "revenue_by_region": revenue_by_region,
        "monthly_trend": monthly_trend,
        "top_customers": top_customers,
        "profit_by_product": profit_by_product
    }