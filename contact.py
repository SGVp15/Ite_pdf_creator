import copy
import os
import re

import pendulum

from Utils.utils import find_numbers_and_ranges, get_dates_certs
from config import OUT_DOCX_PATH, print_docx, OUT_PDF_PATH, map_excel_user, DIR_PDF_FOR_MERGE
from course import Course


class Contact:
    def __init__(self, data: tuple, courses_list: [Course], templates_list: []):
        self.status = False

        self.data = data
        self.files_out_pdf = None
        self.files_out_docx = None

        self.certificate_number = data[0]
        self.date_start = pendulum.instance(data[2])
        self.date_stop = pendulum.instance(data[3])

        if self.date_start is None or self.date_stop is None:
            raise ValueError('date_start')
        self.abr_course = data[4]

        self.issue_date_rus = get_dates_certs(self.date_stop)
        self.course_date_rus = get_dates_certs(self.date_start, self.date_stop)
        self.course_date_eng = get_dates_certs(self.date_start, self.date_stop)
        self.small_date = f'{self.date_start.format('DD.MM.YYYY')}-{self.date_stop.format('DD.MM.YYYY')}'

        self.name_rus = data[7]
        self.name_eng = data[8]
        self.email = data[9]
        if type(self.email) == str:
            self.email = self.email.lower()
        self.gender = data[10]

        del self.data
        if self.abr_course is None:
            raise ValueError(f'{self.certificate_number} abr_course')
        if self.certificate_number is None:
            raise ValueError('sert_number')
        if self.gender is None:
            raise ValueError(f'{self.certificate_number} gender')

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

        self._generate_dir_name()

        self.docx_list_files_name_templates = []
        _a = data[map_excel_user.get('Шаблоны')]
        template_num = find_numbers_and_ranges(_a)
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
        # Удост_MPT_15_октября_2021_Ганц_Павел_32970_aaaa@yan.ru.pdf
        k_print = ''
        temp_path_out_docx = os.path.join(OUT_DOCX_PATH, self.date_start.format('YYYY'), self.date_start.format('MM'),
                                          self.dir_name)
        if file_name in print_docx:
            k_print = 'p_'
            temp_path_out_docx = os.path.join(temp_path_out_docx, DIR_PDF_FOR_MERGE)
        try:
            name = re.sub(r' +', '_', self.name_rus)
        except TypeError:
            pass
        file_out_docx = os.path.join(temp_path_out_docx,
                                     f"{k_print}{file_name[0]}_{self.dir_name}_{self.certificate_number}_{name}"
                                     )
        if self.email:
            file_out_docx += f'_{self.email}'
        file_out_docx += '.docx'

        files_out_pdf = file_out_docx.replace(OUT_DOCX_PATH, OUT_PDF_PATH).replace('.docx', '.pdf')
        return file_out_docx, files_out_pdf

    def __str__(self):
        return f'{self.certificate_number}_{self.abr_course}_{self.small_date}_{self.name_rus}'

    def __eq__(self, other):
        if self.files_out_docx != other.files_out_docx or other.files_out_docx is {}:
            return False
        if (self.certificate_number == other.certificate_number and
                self.name_rus == other.name_rus and
                self.name_eng == other.name_eng and
                self.email == other.email and
                self.gender == other.gender and
                self.date_start == other.date_start and
                self.date_stop == other.date_stop and
                self.docx_list_files_name_templates == other.docx_list_files_name_templates and
                self.course.name_rus == other.course.name_rus and
                self.course.templates == other.course.templates):
            return True

        return False

    def _generate_dir_name(self):
        self.dir_name: str = str(os.path.join(f'{self.abr_course}_{self.small_date}'))
