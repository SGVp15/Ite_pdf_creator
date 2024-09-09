import datetime
import re

from UTILS.utils import replace_month_to_number


def get_date_start_stop_from_strings(s: str) -> (datetime.date, datetime.date):
    s = s.lower()
    s = replace_month_to_number(s)
    s = re.sub(r'\s', '', s)
    try:
        year_start = re.findall(r'\d{4}', s)[0]
        year_stop = re.findall(r'\d{4}', s)[-1]
        s = s.replace(year_start, '')
        s = s.replace(year_stop, '')

        month_start = re.findall(r'\.(\d{2})\.', s)[0]
        month_stop = re.findall(r'\.(\d{2})\.', s)[-1]
        s = s.replace(month_start, '')
        s = s.replace(month_stop, '')

        day_start = re.findall(r'\d+', s)[0]
        day_stop = re.findall(r'\d+', s)[-1]
    except IndexError:
        raise ValueError('date_error')

    try:
        date_start = datetime.date(int(year_start), int(month_start), int(day_start))
        date_stop = datetime.date(int(year_stop), int(month_stop), int(day_stop))
    except ValueError:
        return None, None
    return min(date_start, date_stop), max(date_start, date_stop)
