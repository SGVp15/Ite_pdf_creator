import copy
import os
import re
from datetime import datetime

from UTILS.get_date_start_stop_from_string import get_date_start_stop_from_strings
from UTILS.utils import find_numbers_and_ranges, replace_month_to_number
from config import OUT_DOCX_PATH, print_docx, OUT_PDF_PATH, map_excel_user
from course import Course


class Contact:
    def __init__(self, data: tuple, courses_list: [Course], templates_list: []):
        self.data = data
        self.files_out_pdf = None
        self.files_out_docx = None
        self.date_start: datetime.date = None
        self.date_stop: datetime.date = None
        self.abr_course = self._mapping('AbrCourse')
        self.sert_number = self._mapping('Number')
        self.issue_date_rus = self._mapping('IssueDateRus')
        self.course_date_rus = self._mapping('CourseDateRus')
        self.course_date_eng = self._mapping('CourseDateEng')
        self.name_rus = self._mapping('NameRus')
        self.name_eng = self._mapping('NameEng')
        self.email = self._mapping('Email')
        self.gender = self._mapping('Gender')

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

        self.date_start, self.date_stop = get_date_start_stop_from_strings(self.course_date_rus)
        self.year_start: str = str(self.date_start.year)
        self.month_start: str = str(self.date_start.month)
        self.day_start: str = str(self.date_start.day)
        self.year_stop: str = str(self.date_stop.year)
        self.month_stop: str = str(self.date_stop.month)
        self.day_stop: str = str(self.date_stop.day)

        self.dir_name: str = str(os.path.join(self.year_start, self.date_start.month,
                                              f'{self.year_start}.{self.month_start}.{self.day_start}_{self.abr_course}'))

        self.docx_list_files_name_templates = []
        template_num = find_numbers_and_ranges(data[map_excel_user.get('Template')])
        if len(template_num) == 0:
            self.docx_list_files_name_templates = self.course.templates
        else:
            for i in template_num:
                self.docx_list_files_name_templates.append(templates_list[i])

        self.set_templates(self.docx_list_files_name_templates)

    def _mapping(self, excel):
        return self._clean_str(self.data[map_excel_user.get(excel)])

    @staticmethod
    def _clean_str(s):
        if type(s) is str:
            s = re.sub(r'\s+', ' ', s).strip()
        return s

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

        file_out_docx = os.path.join(
            OUT_DOCX_PATH, self.dir_name,
            f"{k_print}{file_name[0]} {self.date_start.year_start}.{self.date_start.month_start}.{self.date_start.day_start} "
            f"{self.abr_course} {self.sert_number} {self.name_rus}"
        )
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
                self.date_start == other.date_start and
                self.date_stop == other.date_stop and
                self.course_date_rus == other.course_date_rus and
                self.docx_list_files_name_templates == other.docx_list_files_name_templates and
                self.course.name_rus == other.course.name_rus and
                self.course.templates == other.course.templates):
            return True

        return False
