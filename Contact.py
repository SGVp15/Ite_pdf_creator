class Contact:

    def __init__(self):
        self.number: int | None = None
        self.NameRus: str | None = None
        self.NameEng: str | None = None
        self.IssueDateRus: str | None = None

        self.HoursRus: str | None = None
        self.HoursEng: str | None = None

        self.CourseRus: str | None = None
        self.Gender: str | None = None
        self.gender_text: str

        self.email: str | None = None

        self.AbrCourse: str | None = None

        self.Year: str | int | None = None

        self.file_out_docx: str | None = None
        self.file_out_pdf: str | None = None

        self.CourseDateRus: str | None = None
        self.docx_template: str | None = None

        self.NameEng: str | None = None
        self.CourseEng: str | None = None
        self.CourseDateEng: str | None = None
        self.HoursEng: str | None = None

        self.dir_name: str | None = None
        self.course: str | None = None
        self.course_small: str | None = None
        self.lector: str | None = None
