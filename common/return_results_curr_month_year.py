from datetime import datetime


def generate_results_filename():
    """
    Функция генерирует название листа results для текущего месяца и года
    """
    now = datetime.now()
    current_month = now.strftime("%B")
    current_year = now.year
    return f"results_{current_month}_{current_year}"

