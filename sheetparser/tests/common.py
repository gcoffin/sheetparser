import os
import unittest
import six
import numpy as np  # used in client modules (import * from)
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
                         Sequence, StripCellLine, VisibleRows
                         )
from sheetparser.documents import SheetDocument


class LoadWorkbook:
    def setUp(self):
        load_backend(self.backend)
        self.wbk = load_workbook(
            os.path.join(os.path.dirname(__file__), self.filename),
            with_formatting=True)


class TestReadSheetBase(LoadWorkbook):

    def test_empties(self):
        sheet = self.wbk['Sheet1']
        row = six.next(CellRange(sheet, 1, 0, 2, 10).rows())
        row = StripCellLine()(row)
        row = get_value(row)
        self.assertListEqual(row, ['table 1', 'a', 'b', 'c'])

    def test_read(self):
        sheet = self.wbk['Sheet3']
        self.assertEqual(sheet.cell(0, 0).has_borders(BORDERS_VERTICAL), False)
        self.assertEqual(sheet.cell(2, 0).has_borders(BORDERS_VERTICAL), True)

    def test_types(self):
        sheet = self.wbk['Sheet8']
        self.assertEqual(sheet.cell(0, 1).value, datetime.datetime(2017, 12, 1))
        self.assertEqual(sheet.cell(1, 1).value, 1.12)
        self.assertEqual(sheet.cell(2, 1).value, 2.0)
        self.assertEqual(sheet.cell(3, 1).value, datetime.time(2, 30))


class TestFormat(LoadWorkbook):

    date_format = '%Y/%b'

    def test_read_formatted_table(self):
        pattern = Workbook({'Sheet3':
                            Sheet('sheet', Rows,
                                  Line, Empty,
                                  Table(stop=no_horizontal,
                                        table_args=(DEFAULT_TRANSFORMS +
                                                    [RemoveEmptyLines, RemoveEmptyLines('columns')])))
                            })
        context = ListContext()  # PythonObjectContext
        pattern.match_workbook(self.wbk, context)
        result = dict(context.root)
        self.assertEqual(result['table'][0].top_left, [['This']])
        self.assertEqual(result['table'][0].data, [[1, ''], ['', 1], [2, '']])

    def test_merged(self):
        pattern = Workbook({
                'Sheet4': Sheet('sheet', Rows,
                                Empty,
                                Table(table_args=[GetValue, HeaderTableTransform(2), FillData,
                                                  RepeatExisting(0), MergeHeader([0, 1], ch='/'),
                                                  ToDate(0, self.date_format)]))
                })
        context = ListContext()  # PythonObjectContext
        pattern.match_workbook(self.wbk, context)
        result = dict(context.root)
        self.assertEqual(result['table'][0].top_left, [['Year', 'Date']])
        self.assertEqual(result['table'][0].data, [[10, 12, 4], [5, 17, 4]])
        self.assertEqual(result['table'][0].top_headers,
                         [[datetime.datetime(2017, i, 1) for i in [1, 2, 3]]])

    def test_merged2(self):
        pattern = Workbook({
                'Sheet4': Sheet('sheet', Rows,
                                Empty,
                                Table('ignore'),
                                Many(Empty, 2),
                                Table(table_args=[GetValue, HeaderTableTransform(3),
                                                  FillData, RepeatExisting(0),
                                                  MergeHeader([0, 1], ch='/'),
                                                  ToDate(0, self.date_format),
                                                  ToMap]))
                })
        context = ListContext()  # PythonObjectContext
        pattern.match_workbook(self.wbk, context)
        result = dict(context.root)
        expected = {('John', 'Actual', datetime.datetime(2017, 1, 1)): 10,
                    ('John', 'Actual', datetime.datetime(2017, 2, 1)): 12,
                    ('John', 'Forecast', datetime.datetime(2017, 3, 1)): 4,
                    ('Rachel', 'Actual', datetime.datetime(2017, 1, 1)): 5,
                    ('Rachel', 'Actual', datetime.datetime(2017, 2, 1)): 17,
                    ('Rachel', 'Forecast', datetime.datetime(2017, 3, 1)): 4}
        self.assertEqual(result['table'][0].top_left, [['Table 2', 'Date', 'type']])
        self.assertEqual(result['table'][0].data, expected)


class TestSimplePattern(LoadWorkbook):

    def test_pattern1(self):
        pattern = Range('sheet', Rows,
                        Table('t11', table_args=[GetValue, HeaderTableTransform, FillData]))
        range = CellRange(self.wbk['Sheet1'], 1, 1, 5, 5)
        context = ListContext()
        pattern.match_range(range, context)
        self.assertEqual(len(context.root), 2)
        tableresult, = context.root['t11']
        self.assertEqual(tableresult.name, 't11')
        self.assertSequenceEqual(tableresult.top_headers, [['a', 'b', 'c']])
        self.assertSequenceEqual(tableresult.left_headers, [[1, 2, 3]])
        self.assertSequenceEqual(tableresult.data,
                                 [['a11', 'b11', 'c11'],
                                  ['a21', 'b21', 'c21'],
                                  ['a31', 'b31', 'c31']])

    def test_pattern2(self):
        pattern = Range('sheet', Rows, Table('t21', table_args=[GetValue, HeaderTableTransform(1, 1), FillData, RemoveEmptyLines('columns')]))
        range = CellRange(self.wbk['Sheet1'], 1, 1, 5, 10)
        context = ListContext()
        pattern.match_range(range, context)
        self.assertEqual(len(context.root), 2)
        tableresult, = context.root['t21']
        self.assertEqual(tableresult.name, 't21')
        self.assertSequenceEqual(tableresult.top_headers, [['a', 'b', 'c']])
        self.assertSequenceEqual(tableresult.left_headers, [[1, 2, 3]])
        self.assertSequenceEqual(tableresult.data,
                                 [['a11', 'b11', 'c11'],
                                  ['a21', 'b21', 'c21'],
                                  ['a31', 'b31', 'c31']])

    def test_pattern3(self):
        pattern = Sheet('sheet', Rows,
                        Empty,
                        Table('t31', table_args=[GetValue, HeaderTableTransform(1, 2),
                                                 FillData, RemoveEmptyLines('columns')]),
                        Empty,
                        FlexibleRange('t32',
                                      Columns,
                                      Empty,
                                      Table('t33',
                                            table_args=[GetValue, HeaderTableTransform, FillData,
                                                        RemoveEmptyLines('columns')]),
                                      Empty,
                                      Table('t34',
                                            table_args=[GetValue, HeaderTableTransform, FillData,
                                                        Transpose, RemoveEmptyLines])
                                      ))
        context = ListContext()
        pattern.match_range(self.wbk['Sheet1'], context)
        d = dict(context.root)
        self.assertEqual(len(context.root), 4)
        tableresult = d['t31'][0]
        self.assertSequenceEqual(tableresult.top_headers, [['a', 'b', 'c']])
        self.assertSequenceEqual(tableresult.left_headers, [['', '', ''], [1, 2, 3]])
        self.assertSequenceEqual(tableresult.data,
                                 [['a11', 'b11', 'c11'],
                                  ['a21', 'b21', 'c21'],
                                  ['a31', 'b31', 'c31']])
        self.assertSequenceEqual(d['t34'][0].data,
                                 [['a13', 'b13', 'c13'],
                                  ['a23', 'b23', 'c23'],
                                  ['a33', 'b33', 'c33']])

    def test_wbk(self):
        pattern = Workbook({'Sheet1': Sheet('sheet', Rows, Many(Line(name='line') | Empty))})
        context = ListContext()
        pattern.match_workbook(self.wbk, context)

    def test_sequence1(self):
        sheet = self.wbk['Sheet2']
        pattern = Sheet('sheet', Rows,
                        Table, Empty, Table, Empty,
                        Line, Line)
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertEqual(context.table.data[0][0], 'a11')
        self.assertEqual(context.line_1[0], 'line2')

    def test_sequence2(self):
        sheet = self.wbk['Sheet2']
        pattern = Sheet('sheet', Rows,
                        Many((Table('t1',
                                    table_args=[GetValue, HeaderTableTransform,
                                                TableNotEmpty, FillData, ])+Empty) | Line('line')))
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertEqual(context.many[0].add_.t1.data[0][0], 'a11')
        self.assertEqual(context.many[2].line, ['line1'])

    def test_range(self):
        sheet = self.wbk['Sheet1']
        pattern = Range('range', Rows, Table('t1', table_args=[GetValue, FillData]),
                        top=6, bottom=9, left=1, right=4)
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertEqual(context.t1.data,
                         [['table 2', 'a2', 'b2'],
                          [1.0, 'a12', 'b12'],
                          [2.0, 'a22', 'b22'],
                          [3.0, 'a32', 'b32']])

    def test_ranges(self):
        sheet = self.wbk['Sheet1']
        pattern = (Range('range1', Rows,
                         Table('t1',
                               table_args=[GetValue, FillData]),
                         top=6, bottom=9, left=1, right=4) +
                   Range('range2', Rows,
                         Table('t2', table_args=[GetValue, FillData]),
                         top=6, bottom=9, left=1, right=4))

        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertEqual(context.range1.t1.data,
                         [['table 2', 'a2', 'b2'],
                          [1.0, 'a12', 'b12'],
                          [2.0, 'a22', 'b22'],
                          [3.0, 'a32', 'b32']])
        self.assertEqual(context.range2.t2.data,
                         [['table 2', 'a2', 'b2'],
                          [1.0, 'a12', 'b12'],
                          [2.0, 'a22', 'b22'],
                          [3.0, 'a32', 'b32']])


class TestComplex(LoadWorkbook):

    def test_lines(self):
        def has_color(line):
            if not line[0].is_filled:
                raise DoesntMatchException
            return line
        pattern = Sheet('', Columns,
                        Many(Line(line_args=[has_color,
                                             get_value])),
                        Line('after'))
        context = PythonObjectContext()
        pattern.match_range(self.wbk['Sheet5'], context)
        self.assertEquals(len(context.many), 4)
        self.assertEquals(context.after, [2, 4, 6])

    def test_color(self):
        sh = self.wbk['Sheet5']
        self.assertEqual(sh.cell(0, 0).is_filled, True)
        self.assertEqual(sh.cell(0, 1).is_filled, True)
        self.assertEqual(sh.cell(0, 2).is_filled, True)
        self.assertEqual(sh.cell(0, 3).is_filled, True)
        self.assertEqual(sh.cell(0, 4).is_filled, False)

    def test_flexible(self):
        def has_nocolor(line, linecount):
            return not line[0].is_filled
        pattern = Sheet('result', Columns,
                        FlexibleRange('yellow',
                                      Rows, Line,
                                      Table(table_args=[GetValue, HeaderTableTransform(0, 1),
                                                        FillData]),
                                      stop=has_nocolor
                                      ))
        context = ListContext()
        pattern.match_range(self.wbk['Sheet5'], context)
        dct = dict(context.root)
        self.assertEquals(dct['table'][0].data,
                          [[1, 2, 3], [3, 4, 5]])

    def test_complex(self):
        sheet = self.wbk['Sheet6']
        pattern = Sheet('sheet', Columns,
                        Many(Empty),
                        FlexibleRange('f1', Rows,
                                      Many(Empty),
                                      Table('t1',
                                            [GetValue, HeaderTableTransform(2, 1),
                                             FillData, RemoveEmptyLines('columns')],
                                            stop=no_horizontal),
                                      Empty,
                                      FlexibleRange('f2', Columns,
                                                    Many(Empty), Table('t2'),
                                                    stop=no_horizontal),
                                      Many(Empty),
                                      Many((Line('line2',
                                                 [StripCellLine(),
                                                  get_value, Match('Result:', 0)])
                                            + Line('line3', [StripCellLine(), get_value]))
                                           | Line('line1')),
                                      stop=lambda line, linecount: linecount > 2 and empty_line(line)
                                      ),
                        Many(Empty),
                        FlexibleRange('f3', Rows,
                                      Many(Empty),
                                      Table('t3', stop=no_horizontal)))
        context = ListContext()
        pattern.match_range(sheet, context)
        dct = context.root
        self.assertEquals(set(dct.keys()), {'__meta', 't1', 'line2', 'line1', 'line3', 't1', 't2', 't3'})
        self.assertListEqual(dct['line3'], [['End']])
        self.assertEquals(dct['t1'][0].top_left[0][0], 'A more complex example')
        self.assertListEqual(dct['t2'][0].top_left, [['Yet another table']])
        self.assertListEqual(dct['t3'][0].top_left, [['Another table']])
