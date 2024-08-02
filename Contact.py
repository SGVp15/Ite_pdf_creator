import re

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

        # папка по курсам и датам
        dir_name = f"{self.AbrCourse}_{self.CourseDateRus[:-3]}"
        dir_name = dir_name.replace('.', ' ')
        dir_name = dir_name.replace(' ', '')
        self.dir_name = replace_month_to_number(dir_name)

        self.Year = re.findall(r'\d{4}', self.CourseDateRus)[-1]  # замена года выдачиstr
