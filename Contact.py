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

    # папка по курсам и датам
    dir_name = f"{AbrCourse}_{CourseDateRus[:-3]}"
    dir_name = dir_name.replace('.', ' ')
    dir_name = dir_name.replace(' ', '')
    dir_name = replace_month_to_number(dir_name)

    dir_name = dir_name