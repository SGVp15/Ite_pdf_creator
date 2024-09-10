from unittest import TestCase

from UTILS.get_date_start_stop_from_string import get_date_start_stop_from_strings


class Test(TestCase):
    def test_get_date_start_stop_from_strings(self):
        get_date_start_stop_from_strings('22-23 ноября 2018 г.')
        self.fail()
