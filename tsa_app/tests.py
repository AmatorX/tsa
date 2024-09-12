import calendar
import datetime

from django.test import TestCase

from common.days_list import create_date_list


# Create your tests here.


def create_date_list():
    """
    Формируем чанки недель начиная со второго воскресенья года
    """
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

    # Вычисление индекса второго воскресенья
    sundays = [i for i, (_, weekday, _) in enumerate(date_list) if weekday == 'Sun']
    second_sunday_index = sundays[1] if len(sundays) > 1 else None

    # Создание чанков
    chunked_date_list = [date_list[:second_sunday_index]] if second_sunday_index else []
    chunked_date_list.extend([date_list[i:i + 14] for i in range(second_sunday_index, len(date_list), 14)])
    print(f'Чанки дат: {chunked_date_list}')
    for chunk in chunked_date_list:
        print(chunk)
    return chunked_date_list

def get_current_chunk(date_list):
    today = datetime.datetime.now()
    for chunk in date_list:
        start_date_str = f"{chunk[0][0]} {chunk[0][2]} {today.year}"
        end_date_str = f"{chunk[-1][0]} {chunk[-1][2]} {today.year}"
        start_date = datetime.datetime.strptime(start_date_str, "%B %d %Y")
        end_date = datetime.datetime.strptime(end_date_str, "%B %d %Y")

        if start_date <= today <= end_date:
            return chunk
    return None


date_list = create_date_list()
current_chunk = get_current_chunk(date_list)