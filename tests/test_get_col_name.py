import unittest
from excel_tools import get_col_name


class TestGet_col_name(unittest.TestCase):

    def test1(self):
        self.assertEqual(get_col_name(0), "A")

    def test2(self):
        self.assertEqual(get_col_name(1), "B")

    def test3(self):
        self.assertEqual(get_col_name(2), "C")

    def test4(self):
        self.assertEqual(get_col_name(26), "AA")

    def test5(self):
        self.assertEqual(get_col_name(27), "AB")

    def test6(self):
        self.assertEqual(get_col_name(255), "IV")


if __name__ == '__main__':
    unittest.main()
