import json


def save_project(file_path, company, products, customers, sales):
    data = {
        "company": company,
        "products": products,
        "customers": customers,
        "sales": sales
    }

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)


def load_project(file_path):
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except Exception:
        return None