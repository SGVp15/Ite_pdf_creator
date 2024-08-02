import os
import re

from config import TEMPLATES_DIR


def parser_numbers(s: str) -> list:
    s = re.sub(r'[^\d\-]', ' ', s)
    s = re.sub(r'[,\s]*-[,\s]*', '-', s)

    s = s.strip()
    l = s.split()
    clear_int = []
    for x in l:
        try:
            clear_int.append(int(x))
        except:
            la = x.split('-')
            if len(la) > 1:
                la = [int(i) for i in la]
                la.sort()
                la = [i for i in range(la[0], la[-1] + 1)]
                clear_int.extend(la)
    clear_int.sort()
    return clear_int


class Menu:
    def __init__(self):
        self.need_send_email = False

    def get_rows(self, numbers):
        if numbers:
            self.numbers = numbers
            return
        while True:
            print(f'Введите номера строк.\n'
                  f'(Пример: 10-20, 100)')
            answer = input()
            numbers = parser_numbers(answer)
            if len(numbers) > 0:
                print(f'\nВзять эти строки из Excel файла? {answer}\n'
                      f'Y/N')
                answer = input()
                if answer.lower().strip() == 'y':
                    self.numbers = numbers
                    break

    def get_templates(self, templates):
        if templates:
            self.templates = templates
            return
        while True:
            print(f'\nВыберите шаблон:\n'
                  f'(Пример: 1-2, 5)')
            docx_templates = [f for f in os.listdir(TEMPLATES_DIR) if f.endswith('.docx')]
            docx_templates = [f for f in docx_templates if '$' not in f]
            docx_templates.sort()
            for i, docx_template in enumerate(docx_templates):
                print(f'{i}: {docx_template}')
            answer = input()
            try:
                num_templates = parser_numbers(answer)
            except ValueError:
                continue

            if len(num_templates) > 0:
                if min(num_templates) < 0 or max(num_templates) > len(docx_templates):
                    continue
                print(f'\nСоздать по этому шаблону? {num_templates}\n'
                      f'Y/N')
                answer = input()
                if answer.lower().strip() == 'y':
                    self.templates = [docx_templates[i] for i in num_templates if 0 < i <= len(docx_templates)]
                    self.templates = [docx_templates[i] for i in num_templates]
                    break

    def is_need_send_email(self):
        self.need_send_email = False
        while True:
            print(f'\nОтправить по почте?\n'
                  f'Y/N')
            answer = input()
            if answer.lower().strip() == 'y':
                self.need_send_email = True
                break
            elif answer.lower().strip() == 'n':
                self.need_send_email = False
                break
