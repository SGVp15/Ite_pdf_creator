import os

DELETE_DOCX_AFTER_PDF = True
DATA_DIR = os.path.join('//192.168.20.100', 'Administrative server', 'РАБОТА АДМИНИСТРАТОРА', 'ОРГАНИЗАЦИЯ КУРСОВ')
# DATA_DIR = './data'
NAME_OUT_DIR = 'Сертификаты'

OUT_PATH = str(os.path.join(DATA_DIR, NAME_OUT_DIR))

OUT_PDF_PATH = str(os.path.join(OUT_PATH, 'pdf'))
OUT_DOCX_PATH = str(os.path.join(OUT_PATH, 'docx'))

TEMPLATES_DIR = str(os.path.join(DATA_DIR, 'templates'))

os.makedirs(OUT_PDF_PATH, exist_ok=True)
os.makedirs(OUT_DOCX_PATH, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# ---------- EXCEL --------------
FILE_XLSX = '//192.168.20.100/Administrative server/РАБОТА АДМИНИСТРАТОРА/ОРГАНИЗАЦИЯ КУРСОВ/Нумерация с 2015 года.xlsx'
# FILE_XLSX = 'C:/Users/user/PycharmProjects/Ite_pdf_creator/Нумерация с 2015 года.xlsx'
TEMPLATES_DIR = '//192.168.20.100/Administrative server/РАБОТА АДМИНИСТРАТОРА/ОРГАНИЗАЦИЯ КУРСОВ/ШАБЛОНЫ удостоверений'
# TEMPLATES_DIR = 'C:/Users/user/PycharmProjects/Ite_pdf_creator/data/templates'
# FILE_XLSX_TEMP = f'./data/templates/temp.xlsx'
PAGE_NAME = '2015'

map_excel_user = {
    'Number': 0,
    'CourseDateRus': 1,
    'IssueDateRus': 2,
    'CourseDateEng': 3,
    'AbrCourse': 4,
    'Template': 6,
    'NameRus': 7,
    'NameEng': 8,
    'Email': 9,
    'Gender': 10,
    # 'CourseRus': 9,
    # 'CourseEng': 10,
    # 'HoursRus': 11,
    # 'HoursEng': 12,
}
# ('№ сертификата', 'Дата курса (для печати)', 'Дата выдачи (для печати)', 'Дата выдачи сертификата', 'Курс', 'Тренер',
#  'Шаблоны', 'ФИО слушателя', 'ФИО слушателя на латинице', 'Email', 'Пол (М/Ж)', 'Пол по 1 ячейке',
#  'Пол по 2 Ж = 0; M = 1', 'Дата рождения', 'СНИЛС', 'Гражданство код', 'ВО/СПО', 'Фамилия в дипломе', 'Серия диплома',
#  'Номер диплома', 'ID', 'Тренер(рус)', 'Тренер(англ)', 'Должность(рус)', 'Должность(англ)', 'Компания(рус)',
#  'Компания(англ)', 'краткое название', 'название курса на русском', 'название курса на англ.', 'часы на русском',
#  'часы на англ.', 'часы', 'СЖПРОБЕЛЫ ИМЯ (ФИС)', 'Фамилия (ФИС)', 'Имя (ФИС)', 'Отчество (ФИС)',
#  'для пола без отчества (ФИС)', 'Часы(ФИС)', 'Название курса для ФИС', None, None, None)

# --------- TEMPLATES ----------

# Файлы для Печати
print_docx = (
    '0_Удост. на печать.docx',
    '3_Удост. лого Prince2 на печать.docx'
)

log_file = './log.txt'

PICKLE_USERS = './data/users.pk'
PICKLE_FILE_MODIFY = './data/update_file.pk'

_SLEEP_TIME = 60
_LAST_USERS = 100

if os.path.exists(PICKLE_FILE_MODIFY):
    os.remove(PICKLE_FILE_MODIFY)