import unittest

from sheetparser.patterns import VisibleRows
from sheetparser.tests.common import *


class TestReadSheetXLRD(TestReadSheetBase, unittest.TestCase):
    backend = 'sheetparser.backends._xlrd'
    filename = 'test_table1.xls'


class TestReadSheetXLRD_XLSX(TestReadSheetBase, unittest.TestCase):
    backend = 'sheetparser.backends._xlrd'
    filename = 'test_table1.xlsx'


class TestFormatXLRD(TestFormat, unittest.TestCase):
    backend = 'sheetparser.backends._xlrd'
    filename = 'test_table1.xls'
    date_format = '%Y.0/%b'

    def test_fill(self):
        sheet = self.wbk['Sheet7']
        self.assertEqual(sheet.cell(0, 6).value, "{'theme':5}")
        self.assertEqual(sheet.cell(0, 6).fill.type, 'patternFill')
        self.assertEqual(sheet.cell(0, 6).fill.color1, (255, 102, 0))

        self.assertEqual(sheet.cell(0, 7).value, "(200,201,202)")
        self.assertEqual(sheet.cell(0, 7).fill.type, 'patternFill')
        self.assertEqual(sheet.cell(0, 7).fill.color1, (192, 192, 192))

        self.assertEqual(sheet.cell(0, 8).value, "pattern")
        self.assertEqual(sheet.cell(0, 8).fill.type, 'patternFill')
        self.assertEqual(sheet.cell(0, 8).fill.color1, None)

    def test_hidden(self):
        sheet = self.wbk['Sheet7']
        test = [l[0].value for l in VisibleRows().iter_doc(sheet)]
        self.assertEqual(test, ['With hidden rows', '', 'Table 1', 'a1', 'a4'])


class TestSimplePatternXLRD(TestSimplePattern, unittest.TestCase):
    backend = 'sheetparser.backends._xlrd'
    filename = 'test_table1.xls'


class TestSimplePatternXLRD_XLSX(TestSimplePattern, unittest.TestCase):
    backend = 'sheetparser.backends._xlrd'
    filename = 'test_table1.xlsx'


class TestComplexXLRD(TestComplex, unittest.TestCase):
    backend = 'sheetparser.backends._xlrd'
    filename = 'test_table1.xls'


class TestComplexXLRD_XLSX(TestComplex, unittest.TestCase):
    backend = 'sheetparser.backends._xlrd'
    filename = 'test_table1.xlsx'
