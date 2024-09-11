from unittest import TestCase

from UTILS.get_date_start_stop_from_string import get_date_start_stop_from_strings


class Test(TestCase):
    def test_get_date_start_stop_from_strings(self):
        for s in ('09 - 12 сентября 2024 г.', '22-23 ноября 2018 г.','22-23.10.2018 г.'):
            print(s, get_date_start_stop_from_strings(s))
        # self.fail()
