import os
import pickle

from UTILS.log import log
from config import PICKLE_FILE_MODIFY, FILE_XLSX


def get_time_modify_file():
    try:
        info = pickle.load(open(PICKLE_FILE_MODIFY, 'rb'))
    except (FileNotFoundError, IOError):
        return ''
    return info


def check_update_file_excel() -> bool:
    time_file_modify = get_time_modify_file()
    time_file_modify_now = 0
    try:
        time_file_modify_now = os.path.getmtime(FILE_XLSX)
    except (FileNotFoundError, IOError) as e:
        log(e)

    if time_file_modify != time_file_modify_now:
        return True
    else:
        return False