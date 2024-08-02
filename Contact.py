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
        for k, v in from_excel.items():
            pass
