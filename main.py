import sys
import time

from EXCEL.my_excel import read_users_from_excel
from PDF.my_pdf import merge_pdf_contact, create_pdf_contacts
from UTILS.Serialization import load_users, save_users
from UTILS.WORD.my_word import create_docx
from UTILS.files import check_update_file_excel_decorator
from UTILS.log import log
from UTILS.utils import check_config_file, progress
from UTILS.zip import create_zip
from config import _SLEEP_TIME, _LAST_USERS
from contact import Contact
from menu import Menu


def main(is_autorun: bool = False):
    menu = Menu()
    if is_autorun:
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
        create_docx(contact)
        log.info(f'[CREATE_DOCX] {contact.files_out_docx}')
    print('[ OK ]')

    print('CREATE .PDF ... ', end='')
    try:
        create_pdf_contacts(contacts)
    except Exception as e:
        log.error(e)
    print('[ OK ]')

    print('MERGE PDF ... ', end='')
    try:
        merge_pdf_contact(contacts)
    except Exception as e:
        log.error(e)
    print('[ OK ]')

    print('CREATE ZIP ... ', end='')
    try:
        create_zip(contacts)
    except Exception as e:
        log.error(e)
    print('[ OK ]')


@check_update_file_excel_decorator
def auto():
    print('Read Excel')
    all_users_from_file = read_users_from_excel()
    if _LAST_USERS:
        all_users_from_file = all_users_from_file[-min(_LAST_USERS, len(all_users_from_file)):]

    old_users = load_users()

    new_users = [user for user in all_users_from_file if user not in old_users]

    if new_users:
        create_docx_and_pdf(new_users)
        users_for_save = [u for u in new_users if u.status]
        save_users([*users_for_save, *old_users])


if __name__ == '__main__':
    arg = sys.argv
    autorun = False
    if len(arg) > 1:
        if arg[1] in (1, '1'):
            autorun = True
    check_config_file()
    log.info('[ START ]')

    main(autorun)
