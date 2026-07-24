import os
import shutil

from config import FILE_XLSX, TEMPLATES_DIR_SOURCE, TEMPLATES_DIR


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
        r'\s*январ[ья]\s*': '.01.',
        r'\s*феврал[ья]': '.02.',
        r'\s*марта*\s*': '.03.',
        r'\s*апрел[ья]\s*': '.04.',
        r'\s*ма[йя]\s*': '.05.',
        r'\s*июн[ья]\s*': '.06.',
        r'\s*июл[ья]\s*': '.07.',
        r'\s*августа*\s*': '.08.',
        r'\s*сентябр[ья]\s*': '.09.',
        r'\s*октябр[ья]\s*': '.10.',
        r'\s*ноябр[ья]\s*': '.11.',
        r'\s*декабр[ья]\s*': '.12.',

        r's*january ': '.01.',
        r's*february ': '.02.',
        r's*march ': '.03.',
        r's*april ': '.04.',
        r's*may ': '.05.',
        r's*june ': '.06.',
        r's*july ': '.07.',
        r's*august ': '.08.',
        r's*september ': '.09.',
        r's*october ': '.10.',
        r's*november ': '.11.',
        r's*december ': '.12.',

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


import base64
import functools
import hashlib
import re
import sys
from pathlib import Path

from Utils.log import log


def all_exception(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except KeyboardInterrupt:
            print("\nПрограмма принудительно завершена пользователем.")
            sys.exit(0)
        except Exception as e:
            print(f"Произошла ошибка в работе программы: {e}")
            log.error(f'{e}')

    return wrapper


def all_exception_async(func):
    @functools.wraps(func)
    async def wrapper(*args, **kwargs):
        try:
            return await func(*args, **kwargs)
        except KeyboardInterrupt:
            print("\nПрограмма принудительно завершена пользователем.")
            sys.exit(0)
        except Exception as e:
            print(f"Произошла ошибка в работе программы: {e}")
            log.error(f'{e}')

    return wrapper


def to_md5(s: str):
    return hashlib.md5(s.encode()).hexdigest()


def clean_string(s: str) -> str:
    if type(s) is str:
        s = s.replace(',', ', ')
        s = re.sub(r'\s{2,}', ' ', s)
        s = s.strip()
    elif s in ('None', '#N/A', None):
        s = ''
    return s


def file_to_base64(file_path_str: str) -> str:
    path = Path(file_path_str)

    if not path.exists():
        print(f"Ошибка: Путь '{path}' не существует.")
        return None
    if not path.is_file():
        print(f"Ошибка: '{path}' не является файлом.")
        return None

    try:
        file_bytes_data = path.read_bytes()

        # Кодируем в Base64
        base64_bytes = base64.b64encode(file_bytes_data)

        # Возвращаем строку
        return base64_bytes.decode('utf-8')

    except Exception as e:
        print(f"Произошла ошибка при обработке файла: {e}")
        return None


@all_exception
def copy_files():
    for f in os.listdir(TEMPLATES_DIR_SOURCE):
        source = os.path.join(TEMPLATES_DIR_SOURCE, f)
        dist = os.path.join(TEMPLATES_DIR, f)
        if os.path.isfile(source):
            shutil.copy(source, dist)
    # shutil.copy(FILE_XLSX_SOURCE, FILE_XLSX)
