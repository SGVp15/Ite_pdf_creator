import os
import pickle
import sys
import time

import docx2pdf

from EXCEL.my_excel import read_excel_file
from UTILS.files import check_update_file_excel_decorator, delete_empty_folder
from UTILS.log import log
from UTILS.utils import check_config_file, progress
from UTILS.zip import create_zip
from WORD.my_word import create_docx
from config import FILE_XLSX, PICKLE_USERS, DELETE_DOCX_AFTER_PDF, _SLEEP_TIME
from contact import Contact
from course import Course
from menu import Menu


def main(autorun: bool = False):
    menu = Menu()
    if autorun:
        menu.is_auto = 1
    else:
        menu.main()

    if menu.is_auto == 1:
        while True:
            auto()
            for i in range(_SLEEP_TIME):
                progress(text='sleep ', percent=int(i * 100 / _SLEEP_TIME))
                time.sleep(1)
    else:
        rows = menu.get_rows()
        templates_menu = menu.get_templates()

        print('READ EXCEL_FILE ... ', end='')
        users = read_users_from_excel(rows_users=rows)
        for user in users:
            if templates_menu:
                user.set_templates(templates_menu)
        print('[ OK ]')

        create_docx_and_pdf(users)


def create_docx_and_pdf(contacts: [Contact]):
    print('CREATE .DOCX ... ', end='')
    for contact in contacts:
        contact.create_dirs()
        create_docx(contact)
        log.info(f'[CREATE_DOCX] {contact.files_out_docx}')
    print('[ OK ]')

    print('CREATE .PDF ... ', end='')
    for contact in contacts:
        for file_name in contact.docx_list_files_name_templates:
            docx2pdf.convert(contact.files_out_docx[file_name], contact.files_out_pdf[file_name])
            log.info(f'[CREATE_PDF] {contact.files_out_pdf}')
            if DELETE_DOCX_AFTER_PDF:
                path_docx = os.path.dirname(contact.files_out_docx[file_name])
                delete_empty_folder(path_docx)
    print('[ OK ]')

    print('Создаю архив ... ', end='')
    zips = create_zip(contacts)
    print('OK')


@check_update_file_excel_decorator
def auto():
    old_users = []
    new_users = read_users_from_excel()
    new_users = new_users[len(new_users) - 100:len(new_users)]

    try:
        old_users = pickle.load(open(PICKLE_USERS, 'rb'))
    except FileNotFoundError as e:
        log.warning(e)

    new_users = [user for user in new_users if user not in old_users]

    if len(new_users) > 0:
        create_docx_and_pdf(new_users)
        all_users = [*new_users, *old_users]
        pickle.dump(all_users, open(PICKLE_USERS, 'wb'))
        log.info('[Create PICKLE_USERS]')


def read_users_from_excel(file_excel=FILE_XLSX, header=False, rows_users=(-1,)) -> [Contact]:
    data_excel = read_excel_file(file_excel, sheet_names=('2015', 'Курсы', 'Архив Курсов', 'Шаблоны'))
    users_data = data_excel.get('2015')
    users_data = [u for u in users_data if u[1] is not None]
    if rows_users != (-1,):
        users_data = [users_data[i - 1] for i in rows_users]
    elif header is False:
        users_data = users_data[1:]

    courses_data: list = []
    courses_data.extend(data_excel.get('Курсы')[1:])
    courses_data.extend(data_excel.get('Архив Курсов')[1:])

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
        except ValueError:
            log.error(f'[DataError] {data}')
    users = [u for u in users if u.abr_course is not None and u.course is not None]
    return users


if __name__ == '__main__':
    arg = sys.argv
    autorun = False
    if len(arg) > 1:
        if arg[1] in (1, '1'):
            autorun = True
    check_config_file()
    log.info('[ START ]')
    try:
        main(autorun)
    except Exception as e:
        log.error(e)
