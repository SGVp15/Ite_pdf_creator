import re

from config import DOCX_DIR, confirm_docx, print_docx, PDF_DIR
from utils.translit import replace_month_to_number


class Contact:
    def __init__(self, excel):
        self.Number = excel.get('Number')
        self.CourseDateRus = excel.get('CourseDateRus')
        self.IssueDateRus = excel.get('IssueDateRus')
        self.CourseDateEng = excel.get('CourseDateEng')
        self.AbrCourse = excel.get('AbrCourse')
        self.NameRus = excel.get('NameRus')
        self.NameEng = excel.get('NameEng')
        self.Email = excel.get('Email')
        self.Gender = excel.get('Gender')
        self.CourseRus = excel.get('CourseRus')
        self.CourseEng = excel.get('CourseEng')
        self.HoursRus = excel.get('HoursRus')
        self.HoursEng = excel.get('HoursEng')

        self.file_out_pdf = ''
        self.file_out_docx = ''

        dir_name = f"{self.AbrCourse}_{self.CourseDateRus[:-3]}"
        dir_name = re.sub(r'[. ]', '', dir_name)
        self.dir_name = replace_month_to_number(dir_name)

        self.Year = re.findall(r'\d{4}', self.CourseDateRus)[-1]  # замена года выдачиstr
        self.docx_template = ''

    def __call__(self, *args, **kwargs):

        # Удост_MPT_15_октября_2021_Гейнце_Павел_32970_aaa@yandex.ru.pdf
        cert_docx = 'Удост'
        if self.docx_template in confirm_docx:
            cert_docx = 'Подтв'
        k_print = ''
        if self.docx_template in print_docx:
            k_print = 'p_'

        file_out_docx = f"{DOCX_DIR}/{self.dir_name}/{k_print}{cert_docx}_{self.dir_name}_" \
                        f"{self.NameRus}_{self.Number}_{self.Email}.docx"

        file_out_docx = file_out_docx.replace(' ', '_')
        file_out_docx = replace_month_to_number(file_out_docx)
        self.file_out_docx = file_out_docx

        self.file_out_pdf = file_out_docx.replace(DOCX_DIR, PDF_DIR).replace('.docx', '.pdf')
