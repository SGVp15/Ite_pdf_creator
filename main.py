import os
import re
import shutil

import docx2pdf

from EXCEL.my_excel import get_contact_from_excel
from Email import send_email_with_attachment
from WORD.my_word import create_docx
from config import Emails_managers, FILE_XLSX_TEMP, FILE_XLSX
from menu import Menu
from utils import logging
from utils.zip import create_zip


# def convert_all_docx_to_pdf(path_file_out_pdf=f'./output/pdf/', path_file_out_docx='./output/docx/'):
#     docx_files = set([f[:-5] for f in os.listdir(path_file_out_docx) if f.endswith('.docx')])
#     pdf_files = set([f[:-4] for f in os.listdir(path_file_out_pdf) if f.endswith('.pdf')])
#     files_to_convert = docx_files - pdf_files
#     for file in files_to_convert:
#         docx = f'{path_file_out_docx}{file}.docx'
#         pdf = f'{path_file_out_pdf}{file}.pdf'
#         print(f'{docx}  ->  {pdf}')
#         docx2pdf.convert(docx, pdf)


def main():
    menu = Menu()
    menu.get_rows()
    menu.get_templates()
    # menu.is_need_send_email()
    print(f'COPY EXCEL_FILE ... ', end='')
    # shutil.copy2(FILE_XLSX, FILE_XLSX_TEMP)
    print(f'{FILE_XLSX_TEMP} [ OK ]')

    print('READ EXCEL_FILE ... ', end='')
    contacts = get_contact_from_excel(rows_excel=menu.numbers, templates_docx=menu.templates)
    print('[ OK ]')

    print('CREATE .DOCX ... ', end='')
    for contact in contacts:
        create_docx(contact)
    print('[ OK ]')

    print('CREATE .PDF ... ', end='')
    for contact in contacts:
        docx2pdf.convert(contact.file_out_docx, contact.file_out_pdf)
    print('[ OK ]')

    print('Создаю архив ... ', end='')
    zips = create_zip(contacts)
    print('OK')

    if menu.need_send_email:
        print('Отправляю письмо ... ', end='')
        for path_zip in zips:
            # './output/pdf/Agile_28-29.05.2018.zip'
            name = re.sub(r'.*/', '', path_zip)
            name = name.replace('.zip', '')
            send_email_with_attachment(send_to=Emails_managers,
                                       subject=f"Certificates {name}",
                                       text=name,
                                       filename=path_zip)

        print('OK')


if __name__ == '__main__':
    # try:
   main()
    # except Exception as e:
    #     print(e)
    # logging..error(e, exc_info=True)
