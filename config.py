import shutil
from pathlib import Path


def copy_templates_to_local():
    for source in Path(TEMPLATES_DIR_SOURCE).iterdir():
        if source.is_file():
            shutil.copy2(source, Path(TEMPLATES_DIR) / source.name)


IS_DELETE_DOCX_AFTER_CONVERT_PDF = False
OUT_DIR = Path(
    '//192.168.20.100',
    'Administrative server',
    # 'РАБОТА АДМИНИСТРАТОРА',
    # 'ОРГАНИЗАЦИЯ КУРСОВ'
)

NAME_OUT_DIR = 'Удостоверения'
OUT_PATH = Path(OUT_DIR, NAME_OUT_DIR)

OUT_PDF_PATH = Path(OUT_PATH, 'pdf')
OUT_DOCX_PATH = Path(OUT_PATH, 'docx')

# ---------- EXCEL --------------
DATA_DIR = Path(
    '//192.168.20.100',
    'Administrative server',
    'РАБОТА АДМИНИСТРАТОРА',
    'ОРГАНИЗАЦИЯ КУРСОВ'
)
FILE_XLSX_SOURCE = Path(DATA_DIR, 'Нумерация с 2015 года.xlsx')
FILE_XLSX = Path(DATA_DIR, 'Нумерация с 2015 года.xlsx')
TEMPLATES_DIR_SOURCE = Path(DATA_DIR, 'ШАБЛОНЫ удостоверений')
TEMPLATES_DIR = Path('./data/templates')

for dir in (OUT_PDF_PATH, OUT_DOCX_PATH, TEMPLATES_DIR):
    dir.mkdir(parents=True, exist_ok=True)

copy_templates_to_local()

PAGE_NAME = '2015'

list_cells = ['№ сертификата',
              'Номер 1C'
              'ДатаН',
              'ДатаК',
              'Курс',
              'Тренер',
              'Шаблоны',
              'ФИО слушателя',
              'ФИО слушателя на латинице',
              'Email',
              'Пол (М/Ж)', ]

# --------- TEMPLATES ----------

# Файлы для Печати
print_docx = [
    f.name
    for f in Path(TEMPLATES_DIR_SOURCE).iterdir()
    if f.is_file() and 'печать' in f.name
]

FILE_PDF_FOR_MERGE = 'на печать.pdf'
DIR_PDF_FOR_MERGE = 'на печать'

LOG_FILE = './log.txt'

PICKLE_USERS = './data/users.pk'
PICKLE_FILE_MODIFY = './data/update_file.pk'

_SLEEP_TIME = 60
_LAST_USERS = 100

Path(PICKLE_FILE_MODIFY).unlink(missing_ok=True)
