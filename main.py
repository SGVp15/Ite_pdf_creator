import os

import docx2pdf

from EXCEL.my_excel import get_contact_from_excel
from WORD.my_word import create_docx
from config import FILE_XLSX
from menu import Menu
from utils.log import log
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
    if not os.path.exists(FILE_XLSX):
        print(f'file not found [ {FILE_XLSX} ]')
        return

    menu = Menu()
    menu.get_rows()
    menu.get_templates()
    # menu.is_need_send_email()

    # print(f'COPY EXCEL_FILE ... ', end='')
    # shutil.copy2(FILE_XLSX, FILE_XLSX_TEMP)
    # print(f'{FILE_XLSX_TEMP} [ OK ]')

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


if __name__ == '__main__':
    log.info('start')
    try:
        main()
    except Exception as e:
        log.error(e)
