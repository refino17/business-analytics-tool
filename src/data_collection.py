from datetime import datetime

def get_positive_float(prompt):
    while True:
        try:
            value = float(input(prompt))
            if value <= 0:
                print("Value must be positive.")
                continue
            return value
        except:
            print("Invalid number.")

def get_positive_int(prompt):
    while True:
        try:
            value = int(input(prompt))
            if value <= 0:
                print("Value must be positive.")
                continue
            return value
        except:
            print("Invalid integer.")

def get_valid_date(prompt):
    while True:
        date_str = input(prompt)
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return date_str
        except:
            print("Invalid date format. Use YYYY-MM-DD.")

def collect_products():
    products = []
    ids = set()

    n = get_positive_int("Enter number of products: ")

    for i in range(n):
        print(f"\nProduct {i+1}")

        while True:
            pid = input("Product ID: ")
            if pid in ids:
                print("Duplicate Product ID.")
            else:
                ids.add(pid)
                break

        name = input("Product Name: ")
        category = input("Category: ")
        price = get_positive_float("Unit Price: ")

        products.append({
            "Product_ID": pid,
            "Product_Name": name,
            "Category": category,
            "Unit_Price": price
        })

    return products

def collect_customers():
    customers = []
    ids = set()

    n = get_positive_int("\nEnter number of customers: ")

    for i in range(n):
        print(f"\nCustomer {i+1}")

        while True:
            cid = input("Customer ID: ")
            if cid in ids:
                print("Duplicate Customer ID.")
            else:
                ids.add(cid)
                break

        name = input("Customer Name: ")
        city = input("City: ")
        region = input("Region: ").title()

        customers.append({
            "Customer_ID": cid,
            "Customer_Name": name,
            "City": city,
            "Region": region
        })

    return customers

def collect_sales():
    sales = []

    n = get_positive_int("\nEnter number of sales records: ")

    for i in range(n):
        print(f"\nSale {i+1}")

        oid = input("Order ID: ")
        date = get_valid_date("Date (YYYY-MM-DD): ")
        cid = input("Customer ID: ")
        pid = input("Product ID: ")
        qty = get_positive_int("Quantity: ")

        sales.append({
            "Order_ID": oid,
            "Date": date,
            "Customer_ID": cid,
            "Product_ID": pid,
            "Quantity": qty
        })

    return sales