import calendar
import datetime


def create_date_list():
    year = datetime.datetime.now().year
    # Получение имен месяцев и аббревиатур дней недели один раз
    month_names = [calendar.month_name[month] for month in range(1, 13)]
    weekdays_abbr = list(calendar.day_abbr)

    # Создание списка всех дней в году с использованием списковых включений
    date_list = [
        (month_names[month - 1], weekdays_abbr[calendar.weekday(year, month, day)], str(day))
        for month in range(1, 13)
        for day in range(1, calendar.monthrange(year, month)[1] + 1)
]

    # Вычисление индекса первого понедельника
    first_monday_index = next((i for i, (_, weekday, _) in enumerate(date_list) if weekday == 'Mon'), None)

    # Создание чанков
    chunked_date_list = [date_list[:first_monday_index]] if first_monday_index else []
    chunked_date_list.extend([date_list[i:i + 14] for i in range(first_monday_index, len(date_list), 14)])

    return chunked_date_list




# Пример использования и вывода в консоль
# result = create_date_list()
#
# for sublist in result:
#     print(sublist)
#     print('\n')
