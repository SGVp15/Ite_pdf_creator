import copy
import os
import re

from openpyxl import load_workbook

from Contact import Contact
from config import PAGE_NAME, dictory, confirm_docx, print_docx, OUT_DIR, DOCX_DIR, PDF_DIR, FILE_XLSX_TEMP
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
    filename = FILE_XLSX_TEMP
    file_excel = load_workbook(filename=filename, read_only=True, data_only=True)

    contacts_excel = []
    for i in rows_excel:
        from_excel = {}
        for k, v in dictory.items():
            from_excel[k] = read_excel(file_excel, column=v, row=i)
        try:
            contacts_excel.append(Contact(from_excel))
        except TypeError:
            pass
    file_excel.close()

    contacts_return = []

    for contact_excel in contacts_excel:
        for template in templates_docx:
            contact = copy.deepcopy(contact_excel)
            contact.docx_template = template
            contact()
            for _path in ('pdf', 'docx'):
                path_folder = os.path.join(OUT_DIR, _path, contact.dir_name)
                os.makedirs(path_folder, exist_ok=True)

            contacts_return.append(contact)
    return contacts_return
