import os
import re

from openpyxl import load_workbook

from Contact import Contact
from config import PAGE_NAME, FILE_XLSX, dictory, confirm_docx, print_docx, OUT_DIR
from utils.translit import replace_month_to_number


def read_excel(excel, column, row):
    sheet_ranges = excel[PAGE_NAME]
    return str(sheet_ranges[f'{column}{row}'].value)


def clean_export_excel(s):
    s = s.replace(',', ', ')
    # Очистка текста от двойных пробелов
    s = re.sub(r'\s{2,}', ' ', s)
    s = s.strip()
    if s in ('None', '#N/A'):
        s = ''
    return s


def get_contact_from_excel(rows_excel, templates_docx) -> [Contact]:
    # Прочитать excel file -> [Contacts]
    filename = FILE_XLSX
    file_excel = load_workbook(filename=filename, data_only=True)
    contacts = []
    for template in templates_docx:
        for i in rows_excel:
            contact = Contact()

            for k, v in dictory.items():
                contact[k] = clean_export_excel(read_excel(file_excel, column=v, row=i))

            # Создаем папки по курсам
            dir_name = f"{contact.AbrCourse}_{contact.CourseDateRus[:-3]}"
            dir_name = dir_name.replace('.', ' ')
            dir_name = dir_name.replace(' ', '')
            dir_name = replace_month_to_number(dir_name)

            contact.dir_name = dir_name

            for _path in ('pdf', 'docx'):
                path_folder = os.path.join(OUT_DIR, _path, dir_name)
                os.makedirs(path_folder, exist_ok=True)

            # Создаем Объекты по курсам
            contact.docx_template = template

            contact.Year = re.findall(r'\d{4}', contact['CourseDateRus'])[-1]  # замена года выдачи

            # Удост_MPT_15_октября_2021_Гейнце_Павел_32970_aaa@yandex.ru.pdf
            cert_docx = 'Удост'
            if template in confirm_docx:
                cert_docx = 'Подтв'
            k_print = ''
            if template in print_docx:
                k_print = 'p_'

            path = os.path.join(OUT_DIR, 'docx')
            file_out_docx = f"{path}{dir_name}/{k_print}{cert_docx}_{dir_name}_" \
                            f"{contact.last_name_rus}_{contact.first_name_rus}_{contact.number}_{contact.email}.docx"

            # file_out_docx = f"{path}{dir_name}/{k_print}{cert_docx}_{dir_name}_" \
            #                f"{contact['NameRus']}_{contact['Number']}_{template[:-5]}_{contact['Email']}.docx"
            file_out_docx = file_out_docx.replace(' ', '_')
            file_out_docx = replace_month_to_number(file_out_docx)
            contact['file_out_docx'] = file_out_docx

            file_out_pdf = file_out_docx.replace('.docx', '.pdf')
            contact.file_out_pdf = file_out_pdf.replace('/docx/', '/pdf/')
            contacts.append(contact)
    return contacts
