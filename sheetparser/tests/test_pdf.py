import os
import unittest

from sheetparser import (Sheet, Line, PythonObjectContext,
                         load_workbook, Rows,
                         Table, Empty, Workbook
                         )


class TestPdf(unittest.TestCase):
    def test_read_pdf(self):
        filename = os.path.join(os.path.dirname(__file__), 'test_table1.pdf')
        wbk = load_workbook(filename, with_backend='sheetparser.backends._pdfminer')
        pattern = Workbook(
            {2: Sheet('sheet', Rows,
                      Table, Empty, Table, Empty,
                      Line, Line)
             })
        context = PythonObjectContext()
        pattern.match_workbook(wbk, context)
        self.assertEqual(context[0].table.data[0][0], 'a11')
        self.assertEqual(context[0].line_1[0], 'line2')
