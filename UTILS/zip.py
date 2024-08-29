import os.path
import shutil

from config import OUT_PDF_PATH


def create_zip(contacts):
    folders = []
    for contact in contacts:
        folders.append(contact.dir_name)
    folders = set(folders)
    zip_list = []
    for folder in folders:
        shutil.make_archive(os.path.join(OUT_PDF_PATH, str(folder)), 'zip', os.path.join(OUT_PDF_PATH, str(folder)))
        zip_list.append(os.path.join(OUT_PDF_PATH, f'{folder}.zip'))
    return zip_list
