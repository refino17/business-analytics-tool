import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os


def style_chart(ax, title, currency_axis=None):

    ax.set_facecolor("#1E1E1E")
    ax.figure.set_facecolor("#1E1E1E")

    ax.set_title(title, color="white", fontsize=14, weight="bold")

    ax.tick_params(colors="white")

    for spine in ax.spines.values():
        spine.set_color("#444")

    # clean currency axis formatting
    if currency_axis == "y":
        ax.yaxis.set_major_locator(ticker.MaxNLocator(6))
        ax.yaxis.set_major_formatter(
            ticker.FuncFormatter(lambda x, pos: f"₦{x:,.0f}")
        )
    elif currency_axis == "x":
        ax.xaxis.set_major_locator(ticker.MaxNLocator(6))
        ax.xaxis.set_major_formatter(
            ticker.FuncFormatter(lambda x, pos: f"₦{x:,.0f}")
        )


def generate_charts(analysis):

    os.makedirs("charts", exist_ok=True)

    # ================= Revenue by Product =================
    fig, ax = plt.subplots(figsize=(8, 5))

    df = analysis["revenue_by_product"]

    bars = ax.bar(df["Product_Name"], df["Revenue"], color="#2E75B6")

    style_chart(ax, "Revenue by Product", currency_axis="y")

    max_val = df["Revenue"].max()
    ax.set_ylim(0, max_val * 1.2)

    # data labels
    for bar in bars:
        val = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2,
                val + max_val * 0.02,
                f"₦{val:,.0f}",
                ha="center",
                color="white")

    plt.xticks(rotation=40)
    plt.tight_layout()
    plt.savefig("charts/revenue_by_product.png")
    plt.close()

    # ================= Revenue by Region =================
    fig, ax = plt.subplots(figsize=(8, 5))

    df = analysis["revenue_by_region"]

    bars = ax.bar(df["Region"], df["Revenue"], color="#FF8A65")

    style_chart(ax, "Revenue by Region", currency_axis="y")

    max_val = df["Revenue"].max()
    ax.set_ylim(0, max_val * 1.2)

    for bar in bars:
        val = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2,
                val + max_val * 0.02,
                f"₦{val:,.0f}",
                ha="center",
                color="white")

    plt.tight_layout()
    plt.savefig("charts/revenue_by_region.png")
    plt.close()

    # ================= Monthly Trend =================
    fig, ax = plt.subplots(figsize=(8, 5))

    df = analysis["monthly_trend"]

    ax.plot(df["Month"], df["Revenue"],
            marker="o",
            color="#FFC000",
            linewidth=3)

    style_chart(ax, "Monthly Revenue Trend", currency_axis="y")

    max_val = df["Revenue"].max()
    ax.set_ylim(0, max_val * 1.2)

    for x, y in zip(df["Month"], df["Revenue"]):
        ax.text(x, y + max_val * 0.03,
                f"₦{y:,.0f}",
                ha="center",
                color="white")

    plt.tight_layout()
    plt.savefig("charts/monthly_trend.png")
    plt.close()

    # ================= Top Customers =================
    fig, ax = plt.subplots(figsize=(8, 5))

    df = analysis["top_customers"]

    bars = ax.barh(df["Customer_Name"], df["Revenue"], color="#4FC3F7")

    style_chart(ax, "Top Customers", currency_axis="x")

    max_val = df["Revenue"].max()
    ax.set_xlim(0, max_val * 1.25)

    ax.margins(y=0.2)

    for bar in bars:
        val = bar.get_width()
        ax.text(val + max_val * 0.02,
                bar.get_y() + bar.get_height()/2,
                f"₦{val:,.0f}",
                va="center",
                color="white")

    plt.tight_layout()
    plt.savefig("charts/top_customers.png")
    plt.close()

    print("Enterprise BI charts generated successfully.")