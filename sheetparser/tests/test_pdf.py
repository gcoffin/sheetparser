import os
import unittest
import six
import numpy as np
import datetime
from sheetparser import (Document, CellRange, RbColIterator, RbRowIterator,
                         DoesntMatchException, Sheet, Many, Line, PythonObjectContext,
                         load_backend, load_workbook, ResultContext, Columns, Rows,
                         Range, Table, FillData, HeaderTableTransform,
                         RemoveEmptyLines, Empty, FlexibleRange, Transpose,
                         Workbook, BORDERS_VERTICAL, DEFAULT_TRANSFORMS,
                         ListContext, RepeatExisting, MergeHeader, GetValue,
                         ToMap, TableNotEmpty, no_horizontal, ToDate, get_value,
                         Match, empty_line, DebugContext, StripLine,
                         Sequence, StripCellLine
                         )
from sheetparser.documents import SheetDocument


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
