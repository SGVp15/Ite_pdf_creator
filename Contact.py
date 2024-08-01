class Contact:

    def __init__(self):
        self.number: int | None = None
        self.name_eng: str | None = None
        self.last_name_rus: str | None = None
        self.first_name_rus: str | None = None
        self.last_name_eng: str | None = None
        self.first_name_eng: str | None = None

        self.email: str | None = None

        self.Gender: str | None = None
        self.AbrCourse: str | None = None

        self.Year: str | int | None = None

        self.file_out_docx: str | None = None
        self.file_out_pdf: str | None = None

        self.CourseDateRus: str | None = None
        self.docx_template: str | None = None

        self.dir_name: str | None = None
        self.course: str | None = None
        self.course_small: str | None = None
        self.lector: str | None = None
        self.date_from_file = None
        self.date_exam = None
        self.date_exam_connect: str | None = None
        self.remove_at: str | None = None
        self.deadline: str | None = None
        self.scheduled_at: str | None = None
        self.proctor: str | None = None
        self.subject: str | None = None
        self.date_exam_for_subject = None
        self.url_proctor: str | None = None
        self.url_course: str | None = None
        self.id_ispring: str | None = None
        self.status_ispring: str | None = None
        self.identifier: str | None = None
        self.is_create_enrollment: bool = False
        self.status = 'Error'
