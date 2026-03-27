import json

SESSION_FILE = "session_data.json"

def save_session(products, customers, sales):
    data = {
        "products": products,
        "customers": customers,
        "sales": sales
    }

    with open(SESSION_FILE, "w") as f:
        json.dump(data, f)

    print("\nSession saved successfully.")


def load_session():
    try:
        with open(SESSION_FILE, "r") as f:
            data = json.load(f)

        print("\nPrevious session loaded.")
        return data["products"], data["customers"], data["sales"]

    except FileNotFoundError:
        print("\nNo previous session found.")
        return None, None, None