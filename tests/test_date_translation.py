import datetime
import unittest
import xlrd
from excel_tools.exceltools import get_date

class TestDateTranslation(unittest.TestCase):
    def test1(self):
        self.assertEqual(get_date(20151110), datetime.datetime(2015, 11, 10, 0, 0, 0))

    def test2(self):
        self.assertEqual(get_date('2015/11/10'), datetime.datetime(2015, 11, 10, 0, 0, 0))

    def test3(self):
        self.assertEqual(get_date('11/10/2015'), datetime.datetime(2015, 11, 10, 0, 0, 0))

    def test4(self):
        self.assertEqual(get_date('11.10.2015'), datetime.datetime(2015, 11, 10, 0, 0, 0))

    def test5(self):
        self.assertEqual(get_date('20151110'), datetime.datetime(2015, 11, 10, 0, 0, 0))

    def test6(self):
        self.assertEqual(get_date(xlrd.xldate.xldate_from_date_tuple((2015, 11, 10), 0)),
                         datetime.datetime(2015, 11, 10, 0, 0, 0))


if __name__ == '__main__':
    unittest.main()
