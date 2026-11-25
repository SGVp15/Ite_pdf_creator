import unittest

from PDF.my_pdf import merge_all_pdf_in_dir


class MyTestCase(unittest.TestCase):

    def test_merge_all_pdf(self):
        path = (
            '//192.168.20.100/Administrative server/'
            'РАБОТА АДМИНИСТРАТОРА/ОРГАНИЗАЦИЯ КУРСОВ/'
            'Удостоверения/pdf/2026/11/'
            'EAM-online_23.11.2026-26.11.2026/на печать'
        )

        merge_all_pdf_in_dir(path)


if __name__ == '__main__':
    unittest.main()
