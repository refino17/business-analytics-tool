import pandas as pd


def process_data(products, customers, sales):
    df_products = pd.DataFrame(products)
    df_customers = pd.DataFrame(customers)
    df_sales = pd.DataFrame(sales)

    # Ensure optional Cost_Price column exists
    if "Cost_Price" not in df_products.columns:
        df_products["Cost_Price"] = None

    # Convert prices safely
    df_products["Unit_Price"] = pd.to_numeric(df_products["Unit_Price"], errors="coerce")
    df_products["Cost_Price"] = pd.to_numeric(df_products["Cost_Price"], errors="coerce")

    # Convert sales quantity safely
    df_sales["Quantity"] = pd.to_numeric(df_sales["Quantity"], errors="coerce")

    # Merge products
    df_sales = df_sales.merge(
        df_products,
        on="Product_ID",
        how="left"
    )

    # Merge customers
    df_sales = df_sales.merge(
        df_customers,
        on="Customer_ID",
        how="left"
    )

    # Clean region
    if "Region" in df_sales.columns:
        df_sales["Region"] = df_sales["Region"].astype(str).str.title()

    # Revenue
    df_sales["Revenue"] = df_sales["Quantity"] * df_sales["Unit_Price"]

    # Optional cost/profit logic
    df_sales["Cost"] = df_sales["Quantity"] * df_sales["Cost_Price"]

    # If Cost_Price is missing, keep Cost as NaN
    df_sales.loc[df_sales["Cost_Price"].isna(), "Cost"] = pd.NA

    # Profit only where cost exists
    df_sales["Profit"] = df_sales["Revenue"] - df_sales["Cost"]
    df_sales.loc[df_sales["Cost"].isna(), "Profit"] = pd.NA

    # Margin only where cost exists and revenue > 0
    df_sales["Margin"] = (df_sales["Profit"] / df_sales["Revenue"]) * 100
    df_sales.loc[
        df_sales["Cost"].isna() | (df_sales["Revenue"] <= 0),
        "Margin"
    ] = pd.NA

    # Date handling
    df_sales["Date"] = pd.to_datetime(df_sales["Date"], errors="coerce")
    df_sales["Month"] = df_sales["Date"].dt.month_name()
    df_sales["Year"] = df_sales["Date"].dt.year
    df_sales["Month_Num"] = df_sales["Date"].dt.month

    return df_products, df_customers, df_sales