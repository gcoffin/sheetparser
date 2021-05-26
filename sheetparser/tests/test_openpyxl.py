import unittest
import sys
sys.path.append('.')


from sheetparser.patterns import VisibleRows
from sheetparser.tests.common import *


class TestReadSheetOX(TestReadSheetBase, unittest.TestCase):
    backend = 'sheetparser.backends._openpyxl'
    filename = 'test_table1.xlsx'


class TestReadFormatOX(TestFormat, unittest.TestCase):
    backend = 'sheetparser.backends._openpyxl'
    filename = 'test_table1.xlsx'

    def test_fill(self):
        sheet = self.wbk['Sheet7']
        self.assertEqual(sheet.cell(0, 7).value, "(200,201,202)")
        self.assertEqual(sheet.cell(0, 7).fill.type, 'patternFill')
        self.assertEqual(sheet.cell(0, 7).fill.color1, (200, 201, 202))

    def test_fill_theme(self):
        sheet = self.wbk['Sheet7']
        self.assertEqual(sheet.cell(0, 6).value, "{'theme':5}")
        self.assertEqual(sheet.cell(0, 6).fill.type, 'patternFill')
        self.assertEqual(sheet.cell(0, 6).fill.color1, {'theme': 5})

    def test_hidden(self):
        sheet = self.wbk['Sheet7']
        test = [l[0].value for l in VisibleRows().iter_doc(sheet)]
        self.assertEqual(test, ['With hidden rows', '', 'Table 1', 'a1', 'a4', ''])


class TestSimplePatternOX(TestSimplePattern, unittest.TestCase):
    backend = 'sheetparser.backends._openpyxl'
    filename = 'test_table1.xlsx'


class TestComplexOX(TestComplex, unittest.TestCase):
    backend = 'sheetparser.backends._openpyxl'
    filename = 'test_table1.xlsx'

if __name__ == '__main__':
    unittest.main() 