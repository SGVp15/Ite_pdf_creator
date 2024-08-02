from utils.translit import replace_month_to_number


class Contact:

    def __init__(self, from_excel: dict):
        self.Number: str = ''
        self.CourseDateRus: str = ''
        self.IssueDateRus: str = ''
        self.CourseDateEng: str = ''
        self.AbrCourse: str = ''
        self.NameRus: str = ''
        self.NameEng: str = ''
        self.Email: str = ''
        self.Gender: str = ''
        self.CourseRus: str = ''
        self.CourseEng: str = ''
        self.HoursRus: str = ''
        self.HoursEng: str = ''

        # папка по курсам и датам
        dir_name = f"{self.AbrCourse}_{self.CourseDateRus[:-3]}"
        dir_name = dir_name.replace('.', ' ')
        dir_name = dir_name.replace(' ', '')
        dir_name = replace_month_to_number(dir_name)

        self.dir_name = dir_name

        for k, v in from_excel.items():
            pass
