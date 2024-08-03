from datetime import datetime
import os

from dotenv import dotenv_values

DATA_DIR = './data'

OUT_DIR = str(os.path.join(DATA_DIR, 'output'))

PDF_DIR = str(os.path.join(OUT_DIR, 'pdf'))
DOCX_DIR = str(os.path.join(OUT_DIR, 'docx'))

TEMPLATES_DIR = str(os.path.join(DATA_DIR, 'templates'))

os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(DOCX_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

LOG_FILE = './log.txt'

config = dotenv_values('.env')
EMAIL_LOGIN = config['EMAIL_LOGIN']
EMAIL_PASSWORD = config['EMAIL_PASSWORD']

# ---------- EXCEL --------------
FILE_XLSX = '//192.168.20.100/Administrative server/РАБОТА АДМИНИСТРАТОРА/ОРГАНИЗАЦИЯ КУРСОВ/Нумерация с 2015 года.xlsx'

now = datetime.now()  # current date and time

year = now.strftime("%Y")
month = now.strftime("%m")
day = now.strftime("%d")
time = now.strftime("%H:%M:%S")
date_time = now.strftime("%Y-%m-%d-%H-%M-%S")

# FILE_XLSX_TEMP = f'./data/templates/{date_time}.xlsx'
FILE_XLSX_TEMP = f'./data/templates/temp.xlsx'
PAGE_NAME = '2015'

dictory = {
    'Number': 'A',
    'CourseDateRus': 'B',
    'IssueDateRus': 'C',
    'CourseDateEng': 'D',
    'AbrCourse': 'E',
    'NameRus': 'H',
    'NameEng': 'I',
    'Email': 'J',
    'Gender': 'K',
    'CourseRus': 'AC',
    'CourseEng': 'AD',
    'HoursRus': 'AE',
    'HoursEng': 'AF',
}

# --------- TEMPLATES ----------
# Файлы  Подтверждений
confirm_docx = ('ITIL+PRINCE подтверждение.docx',)

# Файлы для Печати
print_docx = ('Удостоверение для печати.docx',
              'ITIL+PRINCE подтверждение.docx',
              'Сертификат для печати.docx',)

# --------- EMAIL ----------
# Куда отправлять Email:
Emails_managers = (
    'p.moiseenko@itexpert.ru',
    'a.katkov@itexpert.ru',
    'a.rybalkin@itexpert.ru',
    #  'g.savushkin@itexpert.ru',
)

SYSTEMLOG = './log.txt'