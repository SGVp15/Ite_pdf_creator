import os
import re

import docx2pdf
from PyPDF2 import PdfMerger

from UTILS.log import log
from config import OUT_DIR_PDF_FOR_PRINT, OUT_PDF_FOR_PRINT, DELETE_DOCX_AFTER_PDF
from contact import Contact


def merge_pdfs(pdf_list, output_file_path):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_file_path)
    merger.close()


def create_pdf_contacts(contacts: [Contact]):
    for contact in contacts:
        for file_name in contact.docx_list_files_name_templates:
            try:
                os.makedirs(os.path.dirname(contact.files_out_pdf[file_name]), exist_ok=True)
                docx2pdf.convert(contact.files_out_docx[file_name], contact.files_out_pdf[file_name])
                log.info(f'[CREATE_PDF] {contact.sert_number} {contact.files_out_pdf}')
                if DELETE_DOCX_AFTER_PDF:
                    os.remove(contact.files_out_docx[file_name])
                    os.removedirs(os.path.dirname(contact.files_out_docx[file_name]))
            except Exception as e:
                log.error(f'[CREATE_PDF] {contact.sert_number} {e}')


def merge_pdf_contact(contacts: [Contact]):
    dirs_pdfs = []
    for contact in contacts:
        for file_name in contact.docx_list_files_name_templates:
            _path = os.path.dirname(contact.files_out_pdf[file_name])
            if re.findall(OUT_DIR_PDF_FOR_PRINT, _path):
                dirs_pdfs.append(_path)
    dirs_pdfs = list(set(dirs_pdfs))

    for dir_pdf in dirs_pdfs:
        pdf_list = [f for f in os.listdir(dir_pdf) if f.endswith(".pdf")]
        try:
            pdf_list.remove(os.path.join(dir_pdf, OUT_PDF_FOR_PRINT))
        except ValueError:
            pass
        pdf_list = [os.path.join(dir_pdf, p) for p in pdf_list]
        out_pdf = os.path.join(dir_pdf, OUT_PDF_FOR_PRINT)
        try:
            os.remove(out_pdf)
        except NotImplementedError:
            pass
        merge_pdfs(pdf_list, out_pdf)
