from datetime import date
from babel.dates import format_date

# print(format_date(today, format='d MMMM YYYY г.', locale='ru'))
# Вывод: 19 ноября 2026 г.

day_1 = date(2021, 11, 3)
day_2 = date(2026, 11, 23)
locale = 'ru'
format = 'd MMMM YYYYг.'
if day_1 == day_2:
    dates_certs = str(format_date(day_1, format=format, locale=locale))
else:
    if day_1.year != day_2.year:
        format_1 = 'dd MMMM YYYYг. - '
        dates_certs = str(format_date(day_1, format=format_1, locale=locale) +
                          format_date(day_2, format=format, locale=locale))
    elif day_1.month != day_2.month:
        format_1 = 'dd MMMM - '
        dates_certs = str(format_date(day_1, format=format_1, locale=locale) +
                          format_date(day_2, format=format, locale=locale))
    elif day_1.month == day_2.month:
        format_1 = 'dd - '
        dates_certs = str(
            format_date(day_1, format=format_1, locale=locale) + format_date(day_2, format=format, locale=locale))
print(dates_certs)
