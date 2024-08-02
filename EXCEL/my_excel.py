import os
import re

from openpyxl import load_workbook

from Contact import Contact
from config import PAGE_NAME, FILE_XLSX, dictory, confirm_docx, print_docx, OUT_DIR, DOCX_DIR, PDF_DIR
from utils.translit import replace_month_to_number


def read_excel(excel, column, row) -> str:
    sheet_ranges = excel[PAGE_NAME]
    return clean_str(str(sheet_ranges[f'{column}{row}'].value))


def clean_str(s):
    s = s.strip()
    s = s.replace(',', ', ')
    s = re.sub(r'\s{2,}', ' ', s)
    if s in ('None', '#N/A'):
        s = ''
    return s


def get_contact_from_excel(rows_excel, templates_docx) -> [Contact]:
    # Прочитать excel file -> [Contacts]
    filename = FILE_XLSX
    file_excel = load_workbook(filename=filename, read_only=True, data_only=True)
    contacts = []

    excel = {}
    for i in rows_excel:
        excel = {}
        for k, v in dictory.items():
            excel[k] = read_excel(file_excel, column=v, row=i)

    for template in templates_docx:
        contact = Contact(excel)

        contact.docx_template = template

        for _path in ('pdf', 'docx'):
            path_folder = os.path.join(OUT_DIR, _path, contact.dir_name)
            os.makedirs(path_folder, exist_ok=True)

        # Удост_MPT_15_октября_2021_Гейнце_Павел_32970_aaa@yandex.ru.pdf
        cert_docx = 'Удост'
        if template in confirm_docx:
            cert_docx = 'Подтв'
        k_print = ''
        if template in print_docx:
            k_print = 'p_'

        file_out_docx = f"{DOCX_DIR}/{contact.dir_name}/{k_print}{cert_docx}_{contact.dir_name}_" \
                        f"{contact.NameRus}_{contact.Number}_{contact.Email}.docx"

        file_out_docx = file_out_docx.replace(' ', '_')
        file_out_docx = replace_month_to_number(file_out_docx)
        contact.file_out_docx = file_out_docx

        contact.file_out_pdf = file_out_docx.replace(DOCX_DIR, PDF_DIR).replace('.docx', '.pdf')

        contacts.append(contact)
    return contacts
