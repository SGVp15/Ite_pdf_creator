import os
import random
import re
import shutil
import time

import docx2pdf
from PyPDF2 import PdfMerger

from UTILS.log import log
from config import DIR_PDF_FOR_MERGE, FILE_PDF_FOR_MERGE, IS_DELETE_DOCX_AFTER_CONVERT_PDF
from contact import Contact


def merge_pdfs(pdf_list, output_file_path):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_file_path)
    print(output_file_path)
    merger.close()


def create_pdf_contacts(contacts: [Contact]):
    for contact in contacts:
        ok_status = []
        for file_name in contact.docx_list_files_name_templates:
            try:
                if not os.path.exists(contact.files_out_pdf[file_name]):
                    # if True:
                    source_doc = contact.files_out_docx[file_name]
                    n = random.randint(111111111, 9111111111)
                    temp_doc = f'./data/{n}.docx'
                    temp_pdf = f'./data/{n}.pdf'
                    dist_pdf = contact.files_out_pdf[file_name]

                    os.makedirs(os.path.dirname(dist_pdf), exist_ok=True)

                    if os.path.isfile(source_doc):
                        shutil.copy(source_doc, temp_doc)

                    docx2pdf.convert(temp_doc, temp_pdf)
                    time.sleep(1)
                    os.remove(temp_doc)
                    if os.path.isfile(temp_pdf):
                        shutil.copy2(temp_pdf, dist_pdf)
                        os.remove(temp_pdf)

                    log.info(f'[CREATE_PDF] {contact.sert_number} {contact.files_out_pdf}')
                    ok_status.append(True)
            except Exception as e:
                log.error(f'[CREATE_PDF] {contact.sert_number} {e}')
                ok_status.append(False)
        if all(ok_status):
            contact.status = True

    if IS_DELETE_DOCX_AFTER_CONVERT_PDF:
        try:
            os.remove(contact.files_out_docx[file_name])
            os.removedirs(os.path.dirname(contact.files_out_docx[file_name]))
        except (NotImplementedError, FileNotFoundError):
            pass


def merge_pdf_contact(contacts: [Contact]):
    dirs_pdfs = []
    for contact in contacts:
        for file_name in contact.docx_list_files_name_templates:
            _path = os.path.dirname(contact.files_out_pdf[file_name])
            if re.findall(DIR_PDF_FOR_MERGE, _path):
                dirs_pdfs.append(_path)

    dirs_pdfs = list(set(dirs_pdfs))

    for dir_pdf in dirs_pdfs:
        pdf_list = [f for f in os.listdir(dir_pdf) if f.endswith(".pdf")]
        try:
            pdf_list.remove(os.path.join(dir_pdf, FILE_PDF_FOR_MERGE))
        except ValueError:
            pass
        pdf_list = [os.path.join(dir_pdf, p) for p in pdf_list]
        out_pdf = os.path.join(dir_pdf, FILE_PDF_FOR_MERGE)
        try:
            os.remove(out_pdf)
        except FileNotFoundError:
            pass
        merge_pdfs(pdf_list, out_pdf)
