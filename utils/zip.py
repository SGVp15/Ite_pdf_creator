import os.path
import shutil

from Contact import Contact
from config import PDF_DIR


def create_zip(contacts: [Contact]):
    folders = []
    for contact in contacts:
        folders.append(contact.dir_name)
    folders = set(folders)
    zip_list = []
    for folder in folders:
        shutil.make_archive(os.path.join(PDF_DIR, str(folder)), 'zip', os.path.join(PDF_DIR, str(folder)))
        zip_list.append(os.path.join(PDF_DIR, f'{folder}.zip'))
    return zip_list
