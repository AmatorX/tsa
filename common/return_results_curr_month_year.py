from datetime import datetime


def generate_results_filename():
    now = datetime.now()
    current_month = now.strftime("%B")
    current_year = now.year
    return f"results_{current_month}_{current_year}"

