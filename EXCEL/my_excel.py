import copy
import os
import re

from openpyxl import load_workbook

from contact import Contact
from config import PAGE_NAME, map_excel_user, OUT_PATH, FILE_XLSX


def read_excel_file(filename=FILE_XLSX, sheet_names=('2015',)) -> {tuple}:
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


def read_excel(excel, column, row) -> str:
    sheet_ranges = excel[PAGE_NAME]
    return clean_str(str(sheet_ranges[f'{column}{row}'].value))


def clean_str(s):
    if type(s) is str:
        s = s.strip()
        s = s.replace(',', ', ')
        s = re.sub(r'\s{2,}', ' ', s)
        if s in ('None', '#N/A'):
            s = ''
    return s


def get_contact_from_excel(rows_excel, templates_docx) -> [Contact]:
    filename = FILE_XLSX
    file_excel = load_workbook(filename=filename, read_only=True, data_only=True)

    contacts_excel = []
    for i in rows_excel:
        from_excel = {}
        for k, v in map_excel_user.items():
            from_excel[k] = read_excel(file_excel, column=v, row=i)
        try:
            contacts_excel.append(Contact(from_excel))
        except TypeError:
            pass
    file_excel.close()

    contacts = []
    for contact_excel in contacts_excel:
        for template in templates_docx:
            contact: Contact = copy.deepcopy(contact_excel)
            contact.docx_list_files_name_templates = template
            contact()
            for _path in ('pdf', 'docx'):
                path_folder = os.path.join(OUT_PATH, _path, contact.dir_name)
                os.makedirs(path_folder, exist_ok=True)

            contacts.append(contact)
    return contacts
