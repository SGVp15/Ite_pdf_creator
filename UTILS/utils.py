import re


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


def parser_numbers(s: str) -> list:
    if s is None:
        return []
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


def replace_month_to_number(s: str):
    # Слоаврь с заменами
    abc = {
        'января': '.01.',
        'февраля': '.02.',
        'марта': '.03.',
        'апреля': '.04.',
        'мая': '.05.',
        'июня': '.06.',
        'июля': '.07.',
        'августа': '.08.',
        'сентября': '.09.',
        'октября': '.10.',
        'ноября': '.11.',
        'декабря': '.12.',
    }

    for key in abc:
        s = s.replace(key, abc[key])
    return s


def progress(text='', percent=0, width=20):
    left = width * percent // 100
    right = width - left
    print(f"\r{text}[{'#' * left}{' ' * right}] {percent:.0f}% ", end='', flush=True)
