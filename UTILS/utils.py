import os
import re

from config import FILE_XLSX


def transliterate(name: str):
    # Слоаврь с заменами
    abc = {'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo',
           'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'i', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n',
           'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h',
           'ц': 'c', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e',
           'ю': 'u', 'я': 'ya', 'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'YO',
           'Ж': 'ZH', 'З': 'Z', 'И': 'I', 'Й': 'I', 'К': 'K', 'Л': 'L', 'М': 'M', 'Н': 'N',
           'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F', 'Х': 'H',
           'Ц': 'C', 'Ч': 'CH', 'Ш': 'SH', 'Щ': 'SCH', 'Ъ': '', 'Ы': 'y', 'Ь': '', 'Э': 'E',
           'Ю': 'U', 'Я': 'YA', ',': '', '?': '', ' ': '_', '~': '', '!': '', '@': '', '#': '',
           '$': '', '%': '', '^': '', '&': '', '*': '', '(': '', ')': '', '-': '', '=': '', '+': '',
           ':': '', ';': '', '<': '', '>': '', '\'': '', '"': '', '\\': '', '/': '', '№': '',
           '[': '', ']': '', '{': '', '}': '', 'ґ': '', 'ї': '', 'є': '', 'Ґ': 'g', 'Ї': 'i',
           'Є': 'e', '—': ''}

    # заменяем все буквы в строке
    for key in abc:
        name = name.replace(key, abc[key])
    return name


def find_numbers_and_ranges(text: str) -> list[int]:
    if type(text) is not str:
        return []
    text = re.sub(r'[^\d\-]', ' ', text)
    text = re.sub(r'[,\s]*-[,\s]*', '-', text)

    text = text.strip()
    list_string = text.split()
    numbers = []
    for x in list_string:
        try:
            numbers.append(int(x))
        except ValueError:
            num = x.split('-')
            if len(num) > 1:
                num = [int(i) for i in num]
                num.sort()
                num = [i for i in range(num[0], num[-1] + 1)]
                numbers.extend(num)
    numbers.sort()
    return numbers


def replace_month_to_number(s: str):
    # Слоаврь с заменами
    abc = {
        r'\s*января\s*': '.01.',
        r'\s*февраля': '.02.',
        r'\s*марта\s*': '.03.',
        r'\s*апреля\s*': '.04.',
        r'\s*мая\s*': '.05.',
        r'\s*июня\s*': '.06.',
        r'\s*июля\s*': '.07.',
        r'\s*августа\s*': '.08.',
        r'\s*сентября\s*': '.09.',
        r'\s*октября\s*': '.10.',
        r'\s*ноября\s*': '.11.',
        r'\s*декабря\s*': '.12.',
    }

    for key in abc:
        s = re.sub(key, abc[key], s)
    return s


def progress(text='', percent=0, width=20):
    left = width * percent // 100
    right = width - left
    print(f"\r{text}[{'#' * left}{' ' * right}] {percent:.0f}% ", end='', flush=True)


def check_config_file():
    if not os.path.exists(FILE_XLSX):
        raise FileNotFoundError
