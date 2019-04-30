# -*- coding: utf-8 -*-
# import csv
import datetime
import decimal
import re

import dateutil.parser
import xlrd
import xlwt
from win32com.client import Dispatch

__author__ = "Jean-Paul Miéville"


class Writer(object):
    """
    This class is used to generate an EXCEL file. The same way as
    csv writer class do.
    """
    re_Movex_Date = re.compile(r'((?:19|20)\d\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])')

    def __init__(self, filename, sheet_name="Output", wrap=False, frozen_headings=True):
        """

        :param filename:
        :param sheet_name:
        :param wrap:
        :param frozen_headings:
        """
        self.filename = filename
        self.book = xlwt.Workbook(encoding='latin-1')
        self.sheet = self.book.add_sheet(sheet_name)
        #
        self.sheet.set_panes_frozen(frozen_headings)
        self.sheet.set_horz_split_pos(1)
        self.sheet.set_remove_splits(frozen_headings)
        #
        self.row_count = 0
        # define the stdandard font
        standard_font = xlwt.Font()
        standard_font.name = 'Calibri'
        standard_font.height = 220
        self.standard_style = xlwt.XFStyle()
        self.standard_style.font = standard_font
        if wrap:
            self.standard_style.alignment.wrap = 1
        else:
            self.standard_style.alignment.wrap = 0
        # Date style
        self.date_style = xlwt.XFStyle()
        self.date_style.font = standard_font
        self._date_format = 'DD.MM.YYYY'
        self.date_style.num_format_str = self._date_format
        if wrap:
            self.date_style.alignment.wrap = 1
        else:
            self.date_style.alignment.wrap = 0
        # font in bold
        bold_font = xlwt.Font()
        bold_font.name = 'Calibri'
        bold_font.height = 220
        bold_font.bold = True
        self.bold_style = xlwt.XFStyle()
        self.bold_style.font = bold_font
        if wrap:
            self.bold_style.alignment.wrap = 1
        else:
            self.bold_style.alignment.wrap = 0
            #

    @property
    def date_format(self):
        return self._date_format

    @date_format.setter
    def date_format(self, date_format):
        self._date_format = date_format
        self.date_style.num_format_str = self._date_format

    def close(self):
        """
        Close the EXCEL File
        """
        self.book.save(self.filename)

    def writerow(self, data, bold=False, excel_date=True):
        """
        This is used to write a row  in the excel file
        if bold parameter is True the full row will be in bold (example total)
        :param data:
        :param bold:
        :param excel_date:
        """
        if bold:
            style = self.bold_style
        else:
            style = self.standard_style
            #
        row = self.sheet.row(self.row_count)
        column_count = 0
        self.row_count += 1
        for d in data:
            if isinstance(d, decimal.Decimal) and excel_date:
                #
                # Handle data coming from the MOVEX where the date is represented by a integer with the following
                # format YYYYMMDD. In this case I translate it to date in EXCEL.
                #
                try:
                    year, month, day = list(map(int, self.re_Movex_Date.match(str(d)).groups()))
                except AttributeError:
                    row.write(column_count, d, style=style)
                else:
                    try:
                        row.write(column_count, datetime.date(year, month, day), style=self.date_style)
                    except ValueError:
                        # The regex found some integer that are not a valid date
                        row.write(column_count, d, style=style)
            elif isinstance(d, (datetime.date, datetime.datetime)):
                # datetime.datetime and datetime.date
                row.write(column_count, d, style=self.date_style)
            else:
                # the cell value is a string. Check if the value is a formula. 
                try:
                    if d.startswith('='):
                        # The value is formula. For the formatting remove the character =.
                        # The character = is use to recognise that the string is formula
                        row.write(column_count, xlwt.Formula(d[1:]), style=style)
                    else:
                        row.write(column_count, d, style=style)
                except (AttributeError, xlwt.ExcelFormulaParser.FormulaParseException):
                    row.write(column_count, d, style=style)
            column_count += 1

    def __enter__(self):
        return self

    def __exit__(self, error, msg, traceback):
        self.close()


class Row(object):
    """
    This class is use to handle the dataModel coming from the Excel sheet.
    The dataModel can then be handled with the column name.
    """

    header = None

    def __init__(self, row, lower=True):
        """

        :param row:
        :param lower:
        :raise:
        """
        self.lower = lower
        for k, cell in zip(Row.header, row):
            if cell.ctype == xlrd.XL_CELL_EMPTY:
                # Empty cell
                v = ''
            elif cell.ctype == xlrd.XL_CELL_TEXT:
                # Cell with text
                v = cell.value
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                # Cell with number
                if cell.value.is_integer():
                    # Test if the number is an integer and then format
                    # the value correctly
                    v = int(cell.value)
                else:
                    # Return a float
                    v = cell.value
            elif cell.ctype == xlrd.XL_CELL_DATE:
                # to be formatted to python date
                v = cell.value
                # *** to be implemented ***
                # datetuple = xlrd.xldate_as_tuple(cell.value, cell.datemode)
                # date = datetime.datetime(datetuple[0], datetuple[1],
                #                          datetuple[2], datetuple[3],
                #                          datetuple[4], datetuple[5])
            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                # Cell with boolean
                v = cell.value
            elif cell.ctype == xlrd.XL_CELL_ERROR:
                # Cell with an error message
                v = xlrd.error_text_from_code[cell.value]
            elif cell.ctype == xlrd.XL_CELL_BLANK:
                # Cell with Blank
                v = ''
            else:
                raise ValueError("Ctype not supported")
                #
            if self.lower:
                self.__dict__[k.lower()] = v
            else:
                self.__dict__[k] = v

    @property
    def data(self):
        if self.lower:
            return [self.__dict__[h.lower()] for h in Row.header]
        else:
            return [self.__dict__[h] for h in Row.header]

    @property
    def string_argument(self):
        """


        :return:
        """
        d = {}
        for k in Row.header:
            if self.lower:
                d.setdefault(k, self.__dict__[k.lower()])
            else:
                d.setdefault(k, self.__dict__[k])
        return d


def clean(value: str) -> str:
    """
    Remove from the column names some forbiden characters

    :param value:
    :return: str
    """
    character_translation = {" ": "_",
                             "#": "",
                             "(": "_",
                             ")": "_",
                             "/": "_",
                             "-": "",
                             "\n": ""}
    #
    if not value:
        # when value is empty the next test will raise an IndexError exception
        return ''
    for character, new_character in list(character_translation.items()):
        value = value.replace(character, new_character)
    value = value.replace('__', '_')
    if value[0] == '_':
        return value[1:]
    elif value[-1] == '_':
        return value[:-1]
    else:
        return value


class ExcelReaderError(Exception):
    pass


def col_name(position):
    """
    Return the column name like EXCEL is using in function of the position
    A, B, C, ..., Z, AA, AB, ..., AZ, BA, ..., BZ, CA, ...

    :param position:
    """
    position = int(position)
    quotient = position / 26
    remainder = position % 26
    if position < 0:
        return ""
    elif position == 0:
        return "A"
    else:
        if quotient < 1:
            return "{}".format(chr(65 + remainder))
        else:
            return "{0}{1}".format(col_name(quotient - 1), chr(65 + remainder))


def reader(excel_file, sheet_name=None, header=True, starting_row=1, lower=True):
    """
    This function open a EXCEL file and return a iterator with the row's content

    :param excel_file:
    :param sheet_name:
    :param header: a logical value indicating whether the file contains the names of the variables as its first
                   line. If missing, the value is determined from the file format: header is set to TRUE if and only
                   if the first row contains one fewer field than the number of columns.
    :param starting_row:
    :param lower: a logical value indicating if we use the lower value of the column name as variable
    """
    workbook = xlrd.open_workbook(excel_file)
    if sheet_name:
        sheet = workbook.sheet_by_name(sheet_name)
    else:
        sheet = workbook.sheet_by_index(0)
    if header:
        # Extract the header
        try:
            Row.header = [clean(col.value) for col in sheet.row(starting_row - 1)]
        except IndexError as error:
            print(sheet.row(starting_row - 1))
            raise error
        if "" in Row.header:
            # Check if the blank are at the end of the list. If they are we will remove them
            index_empty_col = Row.header.index('')
            test = [column for column in Row.header[index_empty_col:] if column]
            if len(test):
                raise ExcelReaderError("One of header value is undefined (col = {0})".format(index_empty_col))
            else:
                Row.header = Row.header[:index_empty_col]
    else:
        header = []
        for i in range(sheet.ncols):
            header.append(col_name(i))
        Row.header = header
        #
    for rx in range(starting_row, sheet.nrows):
        yield Row(sheet.row(rx), lower)


def open_xl_file(file_path, auto_filter=True):
    """
    Open in EXCEL a csv file or xls file
    If autoFilter is True, an automatic filter will be added to
    the first row of the excel sheet.
    """
    #
    xlapp = Dispatch("Excel.Application")
    xlapp.Visible = 1
    xlapp.Workbooks.Open(file_path)
    sheet = xlapp.ActiveSheet
    if auto_filter:
        xlapp.ActiveWindow.SplitRow = 1
        xlapp.ActiveWindow.FreezePanes = True
        # sheet.Rows().AutoFilter()
        sheet.EnableAutoFilter = True
        sheet.Columns("A:IV").AutoFilter(1)
        # Autofit the columns width to the content
    sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, 256)).EntireColumn.AutoFit()
    # xlapp.Save()


class IncorrectDateFormat(Exception):
    pass


def get_date(value):
    """

    :param value:
    :return:
    """
    if isinstance(value, float):
        return datetime.datetime(*xlrd.xldate_as_tuple(value, 0))
    elif isinstance(value, int):
        return dateutil.parser.parse(str(value))
    elif isinstance(value, str):
        try:
            return dateutil.parser.parse(value)
        except ValueError:
            return None
    else:
        raise IncorrectDateFormat("Date format is incorrect ({0}, {1})".format(value, type(value)))


# class ExcelCSV(csv.Dialect):
#     """
#     Describe the usual properties of Excel-generated CSV files.
#     """
#     delimiter = ';'
#     quotechar = '"'
#     # escapechar = None
#     doublequote = True
#     skipinitialspace = False
#     lineterminator = '\r\n'
#     quoting = csv.QUOTE_NONNUMERIC
#
#
# csv.register_dialect("excelCSV", ExcelCSV)


# class CsvWriter(object):
#     """
#     """
#
#     def __init__(self, filename):
#         """
#
#         :param filename:
#         """
#         self.fh = open(filename, 'wb')
#         self.writer = csv.writer(self.fh, dialect=ExcelCSV)
#
#     def close(self):
#         """
#
#
#         """
#         self.fh.close()
#
#     def writerow(self, data):
#         """
#         This is used to write a row  in the excel file
#         if bold parameter is True the full row will be in bold (example total)
#         """
#         self.writer.writerow(data)
#
#     def __enter__(self):
#         """
#
#
#         :return:
#         """
#         return self
#
#     def __exit__(self, error, msg, traceback):
#         self.close()
