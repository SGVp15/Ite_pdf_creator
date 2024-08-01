import shutil


def create_zip(contacts):
    folders = []
    for contact in contacts:
        folders.append(contact['dir_name'])
    folders = set(folders)
    zip_list = []
    for folder in folders:
        shutil.make_archive(f'./output/pdf/{folder}', 'zip', f'./output/pdf/{folder}')
        zip_list.append(f'./output/pdf/{folder}.zip')
    return zip_list
