import copy
import re

from config import OUT_DOCX_DIR, confirm_docx, print_docx, OUT_PDF_DIR, map_excel_user
from course import Course
from UTILS.utils import parser_numbers, replace_month_to_number


class Contact:
    def __init__(self, data, courses_list: [Course], templates_list: []):
        self.abr_course = data[map_excel_user.get('AbrCourse')]
        self.course = None
        for course in courses_list:
            if course.abr == self.abr_course:
                self.course = copy.copy(course)
                self.course_name_rus = self.course.name_rus
                self.course_name_eng = self.course.name_eng
                self.hours_rus = self.course.hour_rus
                self.hours_eng = self.course.hour_eng
                break

        self.sert_number = data[map_excel_user.get('Number')]
        self.course_date_rus = data[map_excel_user.get('CourseDateRus')]
        self.issue_date_rus = data[map_excel_user.get('IssueDateRus')]
        self.course_date_eng = data[map_excel_user.get('CourseDateEng')]
        self.name_rus = data[map_excel_user.get('NameRus')]
        self.name_eng = data[map_excel_user.get('NameEng')]
        self.email = data[map_excel_user.get('Email')]
        self.gender = data[map_excel_user.get('Gender')]

        try:
            dir_name = f"{self.abr_course}_{self.course_date_rus[:-3]}"
        except TypeError:
            pass
        dir_name = re.sub(r'[. ]', '', dir_name)
        self.dir_name = replace_month_to_number(dir_name)

        self.year = re.findall(r'\d{4}', self.course_date_rus)[-1]  # замена года выдачи

        self.docx_list_files_name_templates = []
        template_num = parser_numbers(data[map_excel_user.get('Template')])
        if template_num is None:
            self.docx_list_files_name_templates = self.course.templates
        else:
            for i in template_num:
                self.docx_list_files_name_templates.append(templates_list[i])

        self.set_templates(self.docx_list_files_name_templates)

        if self.abr_course is None or self.abr_course is None or self.course_date_rus is None:
            return

    def set_templates(self, templates_files: list):
        self.docx_list_files_name_templates = templates_files
        self.files_out_pdf = {}
        self.files_out_docx = {}

        for file_name in self.docx_list_files_name_templates:
            self.files_out_docx[file_name], self.files_out_pdf[file_name] = self.create_path_doc_pdf(file_name)

    def __str__(self):
        return f'{self.sert_number} {self.abr_course} {self.course_date_rus} {self.name_rus}'

    def create_path_doc_pdf(self, file_name) -> (str, str):
        # Удост_MPT_15_октября_2021_Гейнце_Павел_32970_aaa@yandex.ru.pdf
        cert_docx = 'Удост'
        if file_name in confirm_docx:
            cert_docx = 'Подтв'
        k_print = ''
        if file_name in print_docx:
            k_print = 'p_'

        file_out_docx = f"{OUT_DOCX_DIR}/{self.dir_name}/{k_print}{cert_docx}_{self.dir_name}_" \
                        f"{self.name_rus}_{self.sert_number}_{self.email}.docx"

        file_out_docx = file_out_docx.replace(' ', '_')
        file_out_docx = replace_month_to_number(file_out_docx)

        files_out_pdf = file_out_docx.replace(OUT_DOCX_DIR, OUT_PDF_DIR).replace('.docx', '.pdf')
        return file_out_docx, files_out_pdf
