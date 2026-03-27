from src.data_collection import collect_products
from src.data_collection import collect_customers
from src.data_collection import collect_sales
from src.session_manager import save_session, load_session
from src.analysis_service import run_full_analysis


def start_new_analysis():

    company_name = input("\nEnter company name: ")

    print("\n--- Data Collection ---\n")

    products = collect_products()
    customers = collect_customers()
    sales = collect_sales()

    save_session(products, customers, sales)

    run_full_analysis(products, customers, sales, company_name)


def load_existing_analysis():

    products, customers, sales = load_session()

    if products is None:
        return

    company_name = input("\nEnter company name for report: ")

    run_full_analysis(products, customers, sales, company_name)