import pendulum


def get_dates_certs(day_1, day_2=None) -> str:
    day_1 = pendulum.instance(day_1)

    if not day_2:
        day_2 = day_1
    else:
        day_2 = pendulum.instance(day_2)

    locale = 'ru'
    format_2 = 'DD MMMM YYYYг.'

    if day_1.year != day_2.year:
        format_1 = 'DD MMMM YYYYг. - '
        dates_certs = str(day_1.format(format_1, locale=locale)) + str(day_2.format(format_2, locale=locale))
    elif day_1.month != day_2.month:
        format_1 = 'DD MMMM - '
        dates_certs = str(day_1.format(format_1, locale=locale)) + str(day_2.format(format_2, locale=locale))
    elif day_1 != day_2:
        format_1 = 'DD - '
        dates_certs = str(day_1.format(format_1, locale=locale)) + str(day_2.format(format_2, locale=locale))
    elif day_1 == day_2:
        dates_certs = str(day_1.format(format_2, locale=locale))

    return dates_certs
