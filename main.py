import os
import pickle

import docx2pdf

from contact import Contact
from EXCEL.my_excel import get_contact_from_excel, read_excel_file
from WORD.my_word import create_docx
from config import FILE_XLSX, pickle_users
from course import Course
from menu import Menu
from UTILS.log import log
from UTILS.zip import create_zip


def main():
    menu = Menu()
    menu.is_auto == 0
    if menu.is_auto == 1:
        contacts = read_users_from_excel()
    else:
        # rows = menu.get_rows()
        rows = (400,)
        # templates = menu.get_templates()
        templates_menu = ('Удостоверение c лого Prince2 пдф.docx',)

        print('READ EXCEL_FILE ... ', end='')
        contacts = read_users_from_excel(rows_users=rows)
        for contact in contacts:
            if templates_menu:
                contact.set_templates(templates_menu)
        print('[ OK ]')

    create_(contacts)


def create_(contacts):
    print('CREATE .DOCX ... ', end='')
    for contact in contacts:
        create_docx(contact)
    print('[ OK ]')

    print('CREATE .PDF ... ', end='')
    for contact in contacts:
        for file_name in contact.docx_list_files_name_templates:
            docx2pdf.convert(contact.files_out_docx[file_name], contact.files_out_pdf[file_name])
    print('[ OK ]')

    print('Создаю архив ... ', end='')
    zips = create_zip(contacts)
    print('OK')


def _aaaA():
    old_users = []
    new_users = []

    new_users = get_contact_from_excel()

    try:
        user: Contact
        old_users = pickle.load(open(pickle_users, 'rb'))
    except FileNotFoundError as e:
        log.error(e)

    new_users = [user for user in new_users if user not in old_users]

    for contact in new_users:
        os.makedirs(os.path.join(OUT_DIR, contact.dir_name), exist_ok=True)

    for i, user in enumerate(new_users):
        os.makedirs(os.path.join(OUT_DIR, user.dir_name), exist_ok=True)
        create_png(user)
        print(f'[{i + 1}/{len(new_users)}]\t{user.file_out_png}')
        log.info(f'[{i + 1}/{len(new_users)}]\t{user.file_out_png}')

    files_cert = []
    for user in new_users:
        files_cert.append(user.file_out_png)

    all_users = [*new_users, *old_users]
    pickle.dump(all_users, open(pickle_users, 'wb'))


def read_users_from_excel(file_excel=FILE_XLSX, header=False, rows_users=(-1,)) -> [Contact]:
    data_excel = read_excel_file(file_excel, sheet_names=('2015', 'Курсы', 'Архив Курсов', 'Шаблоны'))
    users_data = data_excel.get('2015')
    if rows_users != (-1,):
        users_data = [users_data[i - 1] for i in rows_users]
    elif header is False:
        users_data = users_data[:1]

    courses_data: list = []
    courses_data.extend(data_excel.get('Курсы')[1:])
    courses_data.extend(data_excel.get('Архив Курсов')[1:])

    templates_data = data_excel.get('Шаблоны')[1:]

    courses = []
    for c in courses_data:
        courses.append(Course(c))

    templates = []
    for c in templates_data:
        templates.append(c[1])

    users = []
    for data in users_data:
        users.append(Contact(data, courses, templates))
    users = [u for u in users if u.abr_course is not None]
    return users


def check_config_file():
    if not os.path.exists(FILE_XLSX):
        log.error(f'file not found [ {FILE_XLSX} ]')
        raise '[Error] file not found [ {FILE_XLSX} ]'


if __name__ == '__main__':
    check_config_file()
    # log.info('[ START ]')
    # try:
    main()
    # except Exception as e:
    #     log.error(e)
