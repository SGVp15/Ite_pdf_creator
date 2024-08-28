class Course:
    def __init__(self, data: list):
        self.abr = ''  # ['краткое название']
        self.name_rus = ''  # ['название курса на русском']
        self.name_eng = ''  # ['название курса на англ.']
        self.hour_rus = ''  # ['часы на русском']
        self.hour_eng = ''  # ['часы на англ.']
        self.hour = ''  # ['часы']
        self.hour_int = ''  # ['Заполнить']
        self.name_fis = ''  # ['Название для ФИС']
        self.templates = []  # ['Шаблон']

        try:
            self.abr = data[0]
            self.name_rus = data[1]
            self.name_eng = data[2]
            self.hour_rus = data[3]
            self.hour_eng = data[4]
            self.hour = data[5]
            self.hour_int = data[6]
            self.name_fis = data[7]

            for i in [8, 9]:
                if data[i] != '':
                    self.templates.append(data[i])
        except IndexError:
            pass
