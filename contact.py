import copy
import os
import re

import pendulum

from Utils.utils import find_numbers_and_ranges, get_dates_certs
from config import (
    OUT_DOCX_PATH,
    OUT_PDF_PATH,
    DIR_PDF_FOR_MERGE,
    print_docx,
)
from course import Course


class Contact:
    def __init__(
            self,
            data: tuple,
            courses_list: list[Course],
            templates_list: list,
    ):
        self.status = False

        self.data = data
        self.files_out_pdf = None
        self.files_out_docx = None

        # Основные данные
        self.certificate_number = data[0]
        self.date_start = pendulum.instance(data[2])
        self.date_stop = pendulum.instance(data[3])
        self.abr_course = data[4]

        self.name_rus = data[7]
        self.name_eng = data[8]
        self.email = data[9]
        self.gender = data[10]

        if isinstance(self.email, str):
            self.email = self.email.lower()

        # Проверка обязательных полей
        self._validate_data()

        # Даты
        self.issue_date_rus = get_dates_certs(self.date_stop)
        self.course_date_rus = get_dates_certs(
            self.date_start, self.date_stop
        )
        self.course_date_eng = get_dates_certs(
            self.date_start, self.date_stop
        )

        self.small_date = (
            f'{self.date_start.format("DD.MM.YYYY")}-'
            f'{self.date_stop.format("DD.MM.YYYY")}'
        )

        if self.course_date_rus is None:
            raise ValueError(
                f'{self.certificate_number}: course_date_rus'
            )

        # Курс
        self.course = self._get_course(courses_list)

        self._validate_course()

        self.course_name_rus = self.course.name_rus
        self.course_name_eng = self.course.name_eng
        self.hours_rus = self.course.hour_rus
        self.hours_eng = self.course.hour_eng

        # Директория
        self.dir_name = f'{self.abr_course}_{self.small_date}'

        # Выбор шаблонов
        template_num = find_numbers_and_ranges(data[6])

        self.docx_list_files_name_templates = (
            [templates_list[i] for i in template_num]
            if template_num
            else self.course.templates.copy()
        )

        self.set_templates(self.docx_list_files_name_templates)

    def _validate_data(self):
        """Проверка обязательных полей контакта."""

        required_fields = {
            'certificate_number': self.certificate_number,
            'date_start': self.date_start,
            'date_stop': self.date_stop,
            'abr_course': self.abr_course,
            'gender': self.gender,
        }

        empty_fields = [
            field
            for field, value in required_fields.items()
            if value is None or value == ''
        ]

        if empty_fields:
            raise ValueError(
                f'{self.certificate_number}: '
                f'Не заполнены обязательные поля: '
                f'{", ".join(empty_fields)}'
            )

    def _get_course(self, courses_list: list[Course]) -> Course:
        """Поиск и копирование курса."""

        course = next(
            (
                course
                for course in courses_list
                if course.abr == self.abr_course
            ),
            None,
        )

        if course is None:
            raise ValueError(
                f'{self.certificate_number}: course not found '
                f'({self.abr_course})'
            )

        return copy.copy(course)

    def _validate_course(self):
        """Проверка курса."""

        if not self.course.templates:
            raise ValueError(
                f'{self.certificate_number}: templates is null'
            )

    @staticmethod
    def _clean_str(value):
        if isinstance(value, str):
            return re.sub(r'\s+', ' ', value).strip()
        return value

    def set_templates(self, templates_files: list):
        """Формирование путей для выбранных шаблонов."""

        self.docx_list_files_name_templates = templates_files

        self.files_out_pdf = {}
        self.files_out_docx = {}

        for file_name in templates_files:
            docx_path, pdf_path = self.get_path_doc_pdf(file_name)

            self.files_out_docx[file_name] = docx_path
            self.files_out_pdf[file_name] = pdf_path

    def get_path_doc_pdf(self, file_name: str) -> tuple[str, str]:
        """Получение путей DOCX и PDF."""

        name = re.sub(r'\s+', '_', self.name_rus or '')

        # Основная директория
        path = os.path.join(
            OUT_DOCX_PATH,
            self.date_start.format('YYYY'),
            self.date_start.format('MM'),
            self.dir_name,
        )

        # Печатные документы
        prefix = 'p_' if file_name in print_docx else ''

        if prefix:
            path = os.path.join(path, DIR_PDF_FOR_MERGE)

        # Имя файла
        filename = (
            f'{prefix}{file_name[0]}_{self.dir_name}_'
            f'{self.certificate_number}_{name}'
        )

        if self.email:
            filename += f'_{self.email}'

        # DOCX
        docx_path = os.path.join(
            path,
            f'{filename}.docx',
        )

        # PDF
        pdf_path = os.path.join(
            OUT_PDF_PATH,
            os.path.relpath(docx_path, OUT_DOCX_PATH),
        ).replace('.docx', '.pdf')

        return docx_path, pdf_path

    def __str__(self):
        return (
            f'{self.certificate_number}_'
            f'{self.abr_course}_'
            f'{self.small_date}_'
            f'{self.name_rus}'
        )

    def __eq__(self, other):
        if not isinstance(other, Contact):
            return NotImplemented

        if not self.files_out_docx or not other.files_out_docx:
            return False

        if self.files_out_docx != other.files_out_docx:
            return False

        return (
                self.certificate_number == other.certificate_number
                and self.name_rus == other.name_rus
                and self.name_eng == other.name_eng
                and self.email == other.email
                and self.gender == other.gender
                and self.date_start == other.date_start
                and self.date_stop == other.date_stop
                and self.docx_list_files_name_templates == other.docx_list_files_name_templates
                and self.course.name_rus == other.course.name_rus
                and self.course.templates == other.course.templates
        )
