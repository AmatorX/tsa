import calendar
from datetime import datetime

current_date = datetime.today()
# Получаем количество дней в текущем месяце
days_in_month = calendar.monthrange(current_date.year, current_date.month)[1]
print(days_in_month)


