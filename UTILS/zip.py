import os.path
import re
import zipfile
from pathlib import Path

from config import OUT_DIR_PDF_FOR_PRINT
from contact import Contact


def create_zip(contacts: [Contact]):
    dirs_pdfs = []
    for contact in contacts:
        for file_name in contact.docx_list_files_name_templates:
            _path = os.path.dirname(contact.files_out_pdf[file_name])
            if len(re.findall(OUT_DIR_PDF_FOR_PRINT, _path)) == 0:
                dirs_pdfs.append(_path)
    dirs_pdfs = list(set(dirs_pdfs))

    for source_folder in dirs_pdfs:
        dirname = Path(source_folder).name
        archive_name = os.path.join(source_folder, f'{dirname}.zip')
        print(f'{archive_name=}')

        with zipfile.ZipFile(archive_name, 'w') as zipf:
            files = [f for f in os.listdir(source_folder) if f.endswith('.pdf')]
            for file in files:
                # Добавляем файл в архив
                zipf.write(os.path.join(source_folder, file), file)
