import os
import shutil
from pathlib import Path


def copy_templates_to_local():
    for f in os.listdir(TEMPLATES_DIR_SOURCE):
        source = os.path.join(TEMPLATES_DIR_SOURCE, f)
        dist = os.path.join(TEMPLATES_DIR, f)
        if os.path.isfile(source):
            shutil.copy(source, dist)


IS_DELETE_DOCX_AFTER_CONVERT_PDF = False
DATA_DIR = Path('//192.168.20.100', 'Administrative server', 'РАБОТА АДМИНИСТРАТОРА', 'ОРГАНИЗАЦИЯ КУРСОВ')
NAME_OUT_DIR = 'Удостоверения'
OUT_PATH = str(Path(DATA_DIR, NAME_OUT_DIR))

OUT_PDF_PATH = str(os.path.join(OUT_PATH, 'pdf'))
OUT_DOCX_PATH = str(os.path.join(OUT_PATH, 'docx'))

# ---------- EXCEL --------------
FILE_XLSX_SOURCE = Path(DATA_DIR, 'Нумерация с 2015 года.xlsx')
FILE_XLSX = Path(DATA_DIR, 'Нумерация с 2015 года.xlsx')
TEMPLATES_DIR_SOURCE = Path(DATA_DIR, 'ШАБЛОНЫ удостоверений')
TEMPLATES_DIR = './data/templates'

os.makedirs(OUT_PDF_PATH, exist_ok=True)
os.makedirs(OUT_DOCX_PATH, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

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

#  № сертификата 	Номер 1C	ДатаН	ДатаК	Курс	Тренер	Шаблоны	ФИО слушателя	ФИО слушателя на латинице	Email	Пол (М/Ж)
map_excel_user = {}
for i, name in enumerate(list_cells):
    map_excel_user[name] = i

# --------- TEMPLATES ----------

# Файлы для Печати
print_docx = []
for f in os.listdir(TEMPLATES_DIR_SOURCE):
    if 'печать' in f:
        print_docx.append(f)

FILE_PDF_FOR_MERGE = 'на печать.pdf'
DIR_PDF_FOR_MERGE = 'на печать'

LOG_FILE = './log.txt'

PICKLE_USERS = './data/users.pk'
PICKLE_FILE_MODIFY = './data/update_file.pk'

_SLEEP_TIME = 60
_LAST_USERS = 100

if os.path.exists(PICKLE_FILE_MODIFY):
    os.remove(PICKLE_FILE_MODIFY)
