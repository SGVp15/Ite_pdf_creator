from dataclasses import dataclass

from utils.translit import replace_month_to_number


@dataclass
class Contact:
    Number: str
    CourseDateRus: str
    IssueDateRus: str
    CourseDateEng: str
    AbrCourse: str
    NameRus: str
    NameEng: str
    Email: str
    Gender: str
    CourseRus: str
    CourseEng: str
    HoursRus: str
    HoursEng: str

    dir_name: str

    Year: str
