import os

from config import TEMPLATES_DIR
from UTILS.utils import find_numbers_and_ranges


class Menu:
    def __init__(self):
        self.templates = None
        self.numbers = None
        self.is_auto = None
        self.need_send_email = False

    def main(self):
        while True:
            print(f'1: Авто\n'
                  f'0: Ручное')
            self.is_auto = int(input())
            if self.is_auto in (0, 1):
                break

    def get_rows(self):
        while True:
            print(f'Введите номера строк.\n'
                  f'(Пример: 10-20, 100)')
            answer = input()
            numbers = find_numbers_and_ranges(answer)
            if len(numbers) > 0:
                print(f'\nВзять эти строки из Excel файла? {answer}\n'
                      f'Y/N')
                answer = input()
                if answer.lower().strip() == 'y':
                    self.numbers = numbers
                    break
        return self.numbers

    def get_templates(self):
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
                num_templates = find_numbers_and_ranges(answer)
            except ValueError:
                continue

            if len(num_templates) > 0:
                if min(num_templates) < 0 or max(num_templates) > len(docx_templates):
                    continue
                print(f'\nСоздать по этому шаблону? {num_templates}\n'
                      f'Y/N')
                answer = input()
                if answer.lower().strip() == 'y':
                    self.templates = [docx_templates[i] for i in num_templates if 0 <= i < len(docx_templates)]
                    break
        return self.templates

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
