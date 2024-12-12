import pickle

from UTILS.log import log
from config import PICKLE_USERS


def save_users(users):
    with open(PICKLE_USERS, 'wb') as f:
        pickle.dump(users, f)
    log.info('[Create PICKLE_USERS]')


def load_users():
    try:
        users = pickle.load(open(PICKLE_USERS, 'rb'))
        return users
    except FileNotFoundError as e:
        log.warning(e)
        return []
