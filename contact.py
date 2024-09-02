import copy
import os
import re

from UTILS.utils import parser_numbers, replace_month_to_number
from config import OUT_DOCX_PATH, print_docx, OUT_PDF_PATH, map_excel_user
from course import Course


class Contact:
    def __init__(self, data: tuple, courses_list: [Course], templates_list: []):
        self.abr_course = data[map_excel_user.get('AbrCourse')]
        self.sert_number = data[map_excel_user.get('Number')]
        self.issue_date_rus = data[map_excel_user.get('IssueDateRus')]
        self.course_date_rus = data[map_excel_user.get('CourseDateRus')]
        self.course_date_eng = data[map_excel_user.get('CourseDateEng')]
        self.name_rus = data[map_excel_user.get('NameRus')]
        self.name_eng = data[map_excel_user.get('NameEng')]
        self.email = data[map_excel_user.get('Email')]
        self.gender = data[map_excel_user.get('Gender')]

        if self.abr_course is None:
            raise ValueError(f'{self.sert_number} abr_course')
        if self.sert_number is None:
            raise ValueError('sert_number')
        if self.gender is None:
            raise ValueError(f'{self.sert_number} gender')

        self.course = None
        for course in courses_list:
            if course.abr == self.abr_course:
                self.course: Course = copy.copy(course)
                self.course_name_rus = self.course.name_rus
                self.course_name_eng = self.course.name_eng
                self.hours_rus = self.course.hour_rus
                self.hours_eng = self.course.hour_eng
                break
        if self.course is None:
            raise ValueError('course')

        if len(self.course.templates) == 0:
            raise ValueError('templates is null')

        if self.course_date_rus is None:
            raise ValueError('course_date_rus')

        try:
            self.year = re.findall(r'\d{4}', self.course_date_rus)[-1]  # замена года выдачи
            self.month = re.findall(r'\.(\d{2})', replace_month_to_number(self.course_date_rus))[0]
            self.day = re.findall(r'(\d+)\.', replace_month_to_number(self.course_date_rus))[0]
        except IndexError:
            raise ValueError('date_error')

        self.dir_name = os.path.join(self.year, self.month,f'{self.year}.{self.month}.{self.day}_{self.abr_course}')

        self.docx_list_files_name_templates = []
        template_num = parser_numbers(data[map_excel_user.get('Template')])
        if len(template_num) == 0:
            self.docx_list_files_name_templates = self.course.templates
        else:
            for i in template_num:
                self.docx_list_files_name_templates.append(templates_list[i])

        self.set_templates(self.docx_list_files_name_templates)

    def set_templates(self, templates_files: list):
        self.docx_list_files_name_templates = templates_files
        self.files_out_pdf = {}
        self.files_out_docx = {}

        for file_name in self.docx_list_files_name_templates:
            self.files_out_docx[file_name], self.files_out_pdf[file_name] = self.get_path_doc_pdf(file_name)

    def get_path_doc_pdf(self, file_name) -> (str, str):
        # Удост_MPT_15_октября_2021_Гейнце_Павел_32970_aaa@yandex.ru.pdf
        k_print = ''
        if file_name in print_docx:
            k_print = 'p_'

        file_out_docx = os.path.join(OUT_DOCX_PATH, self.dir_name,
                                     f"{k_print}{file_name[0]} {self.year}.{self.month}.{self.day} {self.abr_course} {self.sert_number} {self.name_rus}")
        if self.email:
            file_out_docx += f' {self.email}'
        file_out_docx += '.docx'

        files_out_pdf = file_out_docx.replace(OUT_DOCX_PATH, OUT_PDF_PATH).replace('.docx', '.pdf')
        return file_out_docx, files_out_pdf

    def create_dirs(self):
        os.makedirs(f'{OUT_DOCX_PATH}/{self.dir_name}', exist_ok=True)
        os.makedirs(f'{OUT_PDF_PATH}/{self.dir_name}', exist_ok=True)

    def __str__(self):
        return f'{self.sert_number} {self.abr_course} {self.course_date_rus} {self.name_rus}'

    def __eq__(self, other):
        if self.files_out_docx != other.files_out_docx or other.files_out_docx is {}:
            return False
        if (self.sert_number == other.sert_number and
                self.course_date_rus == other.course_date_rus and
                self.issue_date_rus == other.issue_date_rus and
                self.course_date_eng == other.course_date_eng and
                self.name_rus == other.name_rus and
                self.name_eng == other.name_eng and
                self.email == other.email and
                self.gender == other.gender and
                self.course_date_rus == other.course_date_rus and
                self.docx_list_files_name_templates == other.docx_list_files_name_templates and
                self.course.name_rus == other.course.name_rus and
                self.course.templates == other.course.templates):
            return True

        return False
