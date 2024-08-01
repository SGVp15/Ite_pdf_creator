import os

from dotenv import dotenv_values

OUT_DIR = 'output'

os.makedirs(f'./{OUT_DIR}/pdf', exist_ok=True)
os.makedirs(f'./{OUT_DIR}/docx', exist_ok=True)

TEMPLATES_DIR = './templates'

LOG_FILE = './log.txt'

'''№ сертификата': 'A',
              'Дата курса': 'B',
              'Дата выдачи': 'C',
              'Дата курса англ.': 'D',
              'Курс': 'E',
              'Тренер': 'F',
              'Место проведения': 'G',
              'ФИО слушателя на русском': 'H',
              'ФИО слушателя на латинице': 'I',
              'Пол': 'J',
              'Пол по 1 ячейке': 'K',
              'Пол по 2 Ж = 0; M = 1': 'L',
              'ID': 'M',
              'Тренер(рус)': 'N',
              'Тренер(англ)': 'O',
              'Должность(рус)': 'P',
              'Должность(англ)': 'Q',
              'Компания(рус)': 'R',
              'Компания(англ)': 'S',
              'краткое название': 'T',
              'название курса на русском': 'U',
              'название курса на англ.': 'V',
              'часы на русском': 'W',
              'часы на англ.': 'X',
              'часы': 'Z'}
'''

config = dotenv_values('.env')
EMAIL_LOGIN = config['EMAIL_LOGIN']
EMAIL_PASSWORD = config['EMAIL_PASSWORD']

FILE_XLSX = '//192.168.20.100/Administrative server/РАБОТА АДМИНИСТРАТОРА/ОРГАНИЗАЦИЯ КУРСОВ/Нумерация с 2015 года.xlsx'
# FILE_XLSX = './templates/Нумерация с 2015 года.xlsx'
PAGE_NAME = '2015'

dictory = {
    'NameRus': 'H',
    'NameEng': 'I',
    'CourseRus': 'AC',
    'CourseEng': 'AD',
    'HoursRus': 'AE',
    'HoursEng': 'AF',
    'CourseDateRus': 'B',
    'CourseDateEng': 'D',
    'Email': 'J',
    'Gender': 'K',
    'IssueDateRus': 'C',
    'Number': 'A',
    'AbrCourse': 'E',
}

# Файлы  Подтверждений
confirm_docx = ('ITIL+PRINCE подтверждение.docx',)

# Файлы для Печати
print_docx = ('Удостоверение для печати.docx',
              'ITIL+PRINCE подтверждение.docx',
              'Сертификат для печати.docx',)

# Куда отправлять Email:
Emails_managers = (
    'p.moiseenko@itexpert.ru',
    'a.katkov@itexpert.ru',
    'a.rybalkin@itexpert.ru',
    #  'g.savushkin@itexpert.ru',
)
