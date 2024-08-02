import re

from utils.translit import replace_month_to_number


class Contact:
    def __init__(self, excel):
        self.Number = excel['Number']
        self.CourseDateRus = excel['CourseDateRus']
        self.IssueDateRus = excel['IssueDateRus']
        self.CourseDateEng = excel['CourseDateEng']
        self.AbrCourse = excel['AbrCourse']
        self.NameRus = excel['NameRus']
        self.NameEng = excel['NameEng']
        self.Email = excel['Email']
        self.Gender = excel['Gender']
        self.CourseRus = excel['CourseRus']
        self.CourseEng = excel['CourseEng']
        self.HoursRus = excel['HoursRus']
        self.HoursEng = excel['HoursEng']

        # папка по курсам и датам
        dir_name = f"{self.AbrCourse}_{self.CourseDateRus[:-3]}"
        dir_name = dir_name.replace('.', ' ')
        dir_name = dir_name.replace(' ', '')
        self.dir_name = replace_month_to_number(dir_name)

        self.Year = re.findall(r'\d{4}', self.CourseDateRus)[-1]  # замена года выдачиstr
