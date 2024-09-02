import os
import pickle
import sys
import time

import docx2pdf

from EXCEL.my_excel import read_users_from_excel
from UTILS.files import check_update_file_excel_decorator, delete_empty_folder
from UTILS.log import log
from UTILS.utils import check_config_file, progress
from UTILS.zip import create_zip
from WORD.my_word import create_docx
from config import PICKLE_USERS, DELETE_DOCX_AFTER_PDF, _SLEEP_TIME
from contact import Contact
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
            try:
                docx2pdf.convert(contact.files_out_docx[file_name], contact.files_out_pdf[file_name])
                log.info(f'[CREATE_PDF] {contact.sert_number} {contact.files_out_pdf}')
                if DELETE_DOCX_AFTER_PDF:
                    path_docx = os.path.dirname(contact.files_out_docx[file_name])
                    delete_empty_folder(path_docx)
            except Exception as e:
                log.error(f'[CREATE_PDF] {contact.sert_number} {e}')

    print('[ OK ]')

    print('Создаю архив ... ', end='')
    zips = create_zip(contacts)
    print('OK')


@check_update_file_excel_decorator
def auto():
    old_users = []
    print('Read Excel')
    new_users = read_users_from_excel()

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
