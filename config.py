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
# FILE_XLSX = '//192.168.20.100/Administrative server/РАБОТА АДМИНИСТРАТОРА/ОРГАНИЗАЦИЯ КУРСОВ/Нумерация с 2015 года.xlsx'
FILE_XLSX = './data/templates/Нумерация с 2015 года.xlsx'
PAGE_NAME = '2015'


class dictory:
    def init(self):
        self.Number: str = 'A'
        self.CourseDateRus: str = 'B'
        self.IssueDateRus: str = 'C'
        self.CourseDateEng: str = 'D'
        self.AbrCourse: str = 'E'
        self.NameRus: str = 'H'
        self.NameEng: str = 'I'
        self.Email: str = 'J'
        self.Gender: str = 'K'
        self.CourseRus: str = 'AC'
        self.CourseEng: str = 'AD'
        self.HoursRus: str = 'AE'
        self.HoursEng: str = 'AF'

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
