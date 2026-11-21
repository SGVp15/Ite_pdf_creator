import re

from openpyxl import load_workbook

from config import PAGE_NAME, FILE_XLSX
from contact import Contact
from course import Course


def read_excel_file(filename=FILE_XLSX, sheet_names=('2015',)) -> dict:
    workbook = load_workbook(filename=filename, read_only=True, data_only=True)
    all_data = {}
    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        data = []
        for row in sheet.iter_rows(values_only=True):
            # data.append(list(map(clean_str, row)))
            data.append(row)
        all_data[sheet_name] = data
    workbook.close()
    return all_data


def read_users_from_excel(file_excel=FILE_XLSX, header=False, rows_users=(),
                          sheet_names=('2015', 'Курсы', 'Шаблоны')) -> [Contact]:
    data_excel = read_excel_file(file_excel, sheet_names=sheet_names)
    users_data = data_excel.get('2015')
    if rows_users:
        users_data = [users_data[i - 1] for i in rows_users]
    elif header is False:
        users_data = users_data[1:]

    users_data = [u for u in users_data if u[0] is not None]

    courses_data: list = []
    courses_data.extend(data_excel.get('Курсы')[1:])

    templates_data = data_excel.get('Шаблоны')[1:]

    courses = []
    for c in courses_data:
        try:
            courses.append(Course(c))
        except ValueError:
            pass

    templates = []
    for t in templates_data:
        templates.append(t[1])

    users = []
    for data in users_data:
        try:
            users.append(Contact(data, courses, templates))
        except Exception as e:
            pass

    users = [u for u in users if u.abr_course is not None and u.course is not None]
    return users
