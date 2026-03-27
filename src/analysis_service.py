from src.data_processing import process_data
from src.analysis import generate_analysis
from src.charts import generate_charts
from src.excel_export import export_to_excel


def run_full_analysis(products, customers, sales, company_name, excel_output_file):

    df_products, df_customers, df_sales = process_data(products, customers, sales)

    analysis = generate_analysis(df_sales)

    generate_charts(analysis)

    export_to_excel(
        df_products,
        df_customers,
        df_sales,
        analysis,
        company_name,
        excel_output_file
    )

    return True