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
                         Workbook, BORDERS_VERTICAL, DEFAULT_TRANFORMS,
                         ListContext, RepeatExisting, MergeHeader, GetValue,
                         ToMap, TableNotEmpty, no_horizontal, ToDate, get_value,
                         Match, empty_line, DebugContext, StripLine
                         )
from sheetparser.documents import SheetDocument


class DummyWorkbook(Document):
    def __init__(self, sheets):
        self.sheets = sheets

    def __iter__(self):
        for name, w in six.iteritems(self.sheets):
            yield DummySheet(name, w)


class DummySheet(SheetDocument, CellRange):
    def __init__(self, name, array):
        self.name = name
        self.rge = np.array(array)
        self.left = self.top = 0
        self.bottom, self.right= self.rge.shape

    def is_hidden(self):
        return False

    def abscell(self, row, col):
        return DummyCell(self.rge[row, col])

    cell = abscell #NO!!

    def __repr__(self):
        return "<DummySheet %s>"%self.rge


class DummyCell(object):
    def __init__(self, value):
        self.value = value
        self.is_merged = False

    def is_empty(self):
        return self.value == 0


def to_list_value(l):
    return [i.value for i in l]


class TestColIterator(unittest.TestCase):
    def test_row_iters(self):
        test_array = [[0, 0, 1, 1, 0]]*3
        sheet = DummySheet('test', test_array)
        it = RbColIterator(sheet)
        for col in zip(*test_array):
            self.assertSequenceEqual(to_list_value(six.next(it)), list(col))
        it = RbRowIterator(sheet)
        for row in test_array:
            self.assertSequenceEqual(to_list_value(six.next(it)), list(row))

    def test_rowiter_rollback(self):
        test_array = [[0, 0, 1, 1, 0]]*3
        sheet = DummySheet('test', test_array)
        it = RbRowIterator(sheet)
        six.next(it)
        with it.rollback_if_fail(reraise=False):
            for row in test_array[1:]:
                self.assertSequenceEqual(to_list_value(six.next(it)), list(row))
            raise DoesntMatchException()
        for row in test_array[1:]:
            self.assertSequenceEqual(to_list_value(six.next(it)), list(row))

    def test_subrange(self):
        test_array = np.arange(20, dtype=int).reshape(4, 5)
        sheet = DummySheet('test', test_array)
        r = CellRange(sheet, 0, 0, 1, 1) #top left is 7, bottom right is 13
        it = RbColIterator(r)
        sub_array = test_array[0:1, 0:1]
        for col in zip(*sub_array):
            self.assertSequenceEqual(to_list_value(six.next(it)), list(col))

    # 0  1  2  3  4
    # 5  6  7  8  9
    # 10 11 12 13 14
    # 15 16 17 18 19
    def test_subrange2(self):
        test_array = np.arange(20, dtype=int).reshape(4, 5)
        sheet = DummySheet('test', test_array)
        r = CellRange(sheet, 1, 2, 3, 4) #top left is 7, bottom right is 13
        it = RbColIterator(r)
        sub_array = test_array[1:3, 2:4]
        for col in zip(*sub_array):
            self.assertSequenceEqual(to_list_value(six.next(it)), list(col))
        it = RbRowIterator(r)
        for row in sub_array:
            self.assertSequenceEqual(to_list_value(six.next(it)), list(row))

    def test_subsubrange(self):
        test_array = np.arange(20, dtype=int).reshape(4, 5)
        sheet = DummySheet('test', test_array)
        r = CellRange(sheet, 1, 1, 5, 5) #top left is 7, bottom right is 13
        sr = CellRange(r, 1, 1, 2, 2)
        it = RbRowIterator(sr)
        for row in test_array[2:3, 2:3]:
            self.assertSequenceEqual(to_list_value(six.next(it)), list(row))


class TestArray(unittest.TestCase):
    def test_rollback(self):
        test_array = np.array([[1]*5])
        sheet = DummySheet('test', test_array)
        pattern = Sheet('result', Rows,
                        Many(Line, min=2)|Line('line'))
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertSequenceEqual(context.root['line'], [1, 1, 1, 1, 1])


class TestSimpleExcel(unittest.TestCase):
    def setUp(self):
        load_backend('sheetparser.backends._xlrd')
        self.wbk = load_workbook(os.path.join(os.path.dirname(__file__), 'test_table1.xlsx'), with_formatting=False)
        self.sheet = self.wbk['Sheet1']

    def test_pattern1(self):
        pattern = Range('sheet', Rows,
                        Table('t11', table_args=[GetValue, HeaderTableTransform, FillData]))
        range = CellRange(self.sheet, 1, 1, 5, 5)
        context = ListContext()
        pattern.match_range(range, context)
        self.assertEqual(len(context.root), 2)
        tableresult, = context.root['t11']
        self.assertEqual(tableresult.name, 't11')
        self.assertSequenceEqual(tableresult.top_headers, [['a', 'b', 'c']])
        self.assertSequenceEqual(tableresult.left_headers, [[1, 2, 3]])
        self.assertSequenceEqual(tableresult.data, [['a11', 'b11', 'c11'], ['a21', 'b21', 'c21'], ['a31', 'b31', 'c31']])

    def test_pattern2(self):
        pattern = Range('sheet', Rows, Table('t21', table_args=[GetValue, HeaderTableTransform(1, 1), FillData, RemoveEmptyLines('columns')]))
        range = CellRange(self.sheet, 1, 1, 5, 10)
        context = ListContext()
        pattern.match_range(range, context)
        self.assertEqual(len(context.root), 2)
        tableresult, = context.root['t21']
        self.assertEqual(tableresult.name, 't21')
        self.assertSequenceEqual(tableresult.top_headers, [['a', 'b', 'c']])
        self.assertSequenceEqual(tableresult.left_headers, [[1, 2, 3]])
        self.assertSequenceEqual(tableresult.data, [['a11', 'b11', 'c11'], ['a21', 'b21', 'c21'], ['a31', 'b31', 'c31']])

    def test_pattern3(self):
        pattern = Sheet('sheet', Rows,
                        Empty,
                        Table('t31', table_args=[GetValue, HeaderTableTransform(1, 2),
                                                 FillData, RemoveEmptyLines('columns')]),
                        Empty,
                        FlexibleRange('t32',
                                      Columns,
                                      Empty,
                                      Table('t33', table_args=[GetValue, HeaderTableTransform, FillData,
                                                               RemoveEmptyLines('columns')]),
                                      Empty,
                                      Table('t34', table_args=[GetValue, HeaderTableTransform, FillData,
                                                               Transpose, RemoveEmptyLines])
                                      ))
        context = ListContext()
        pattern.match_range(self.sheet, context)
        d = dict(context.root)
        self.assertEqual(len(context.root), 4)
        tableresult = d['t31'][0]
        self.assertSequenceEqual(tableresult.top_headers, [['a', 'b', 'c']])
        self.assertSequenceEqual(tableresult.left_headers, [['', '', ''], [1, 2, 3]])
        self.assertSequenceEqual(tableresult.data, [['a11', 'b11', 'c11'], ['a21', 'b21', 'c21'], ['a31', 'b31', 'c31']])
        self.assertSequenceEqual(d['t34'][0].data, [['a13', 'b13', 'c13'], ['a23', 'b23', 'c23'], ['a33', 'b33', 'c33']])

    def test_wbk(self):
        pattern = Workbook({'Sheet1': Sheet('sheet', Rows, Many(Line(name='line')|Empty))})
        context = ListContext()
        pattern.match_workbook(self.wbk, context)


    def test_sequence1(self):
        sheet = self.wbk['Sheet2']
        pattern = Sheet('sheet', Rows,
                        Table, Empty, Table, Empty,
                        Line, Line
                        )
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertEqual(context.table.data[0][0], 'a11')
        self.assertEqual(context.line_1[0], 'line2')

    def test_sequence2(self):
        sheet = self.wbk['Sheet2']
        pattern = Sheet('sheet', Rows,
                       Many( (Table('t1', table_args=[GetValue, HeaderTableTransform, 
                                                      TableNotEmpty, FillData, ])+Empty) | Line('line'))
                        )
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertEqual(context.many[0].t1.data[0][0], 'a11')
        self.assertEqual(context.many[2][0], 'line1')


class TestOpenpyxl(unittest.TestCase):
    def setUp(self):
        load_backend('sheetparser.backends._openpyxl')
        self.wbk = load_workbook(
            os.path.join(os.path.dirname(__file__),
                         'test_table1.xlsx'),
            with_formatting=True)

    def test_read(self):
        sheet = self.wbk['Sheet3']
        self.assertEqual( sheet.cell(0, 0).has_borders(BORDERS_VERTICAL), False)
        self.assertEqual( sheet.cell(2, 0).has_borders(BORDERS_VERTICAL), True)

    def test_read_formatted_table(self):
        pattern = Workbook({'Sheet3':
                            Sheet('sheet', Rows,
                                  Line, Empty, 
                                  Table(stop=no_horizontal, 
                                        table_args=DEFAULT_TRANFORMS+[RemoveEmptyLines, RemoveEmptyLines('columns')]))
                            })
        context = ListContext() #PythonObjectContext
        pattern.match_workbook(self.wbk, context)
        result = dict(context.root)
        self.assertEqual(result['table'][0].top_left, [['This']])
        self.assertEqual(result['table'][0].data, [[1, ''], ['', 1], [2, '']])


    def test_merged(self):
        pattern = Workbook({'Sheet4':
                                Sheet('sheet', Rows,
                                      Empty, Table(table_args=[GetValue, HeaderTableTransform(2), FillData,
                                                               RepeatExisting, MergeHeader([0, 1], ch='/'), ToDate(0, '%Y/%b')]))
                            })
        context = ListContext() #PythonObjectContext
        pattern.match_workbook(self.wbk, context)
        result = dict(context.root)
        self.assertEqual(result['table'][0].top_left, [['Year', 'Date']])
        self.assertEqual(result['table'][0].data, [[10, 12, 4], [5, 17, 4]])
        self.assertEqual(result['table'][0].top_headers, [[datetime.datetime(2017, i, 1) for i in [1, 2, 3]]])


    def test_merged2(self):
        pattern = Workbook({'Sheet4':
                            Sheet('sheet', Rows,
                                  Empty,
                                  Table('ignore'),
                                  Many(Empty, 2),
                                  Table(table_args=[GetValue, HeaderTableTransform(3), FillData,
                                                    RepeatExisting, MergeHeader([0, 1], ch='/'), ToDate(0, '%Y/%b'),
                                                    ToMap]))
                            })
        context = ListContext() #PythonObjectContext
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


class TestWin32(unittest.TestCase):
    def setUp(self):
        load_backend('sheetparser.backends._win32com')
        self.wbk = load_workbook(os.path.join(os.path.dirname(__file__), 'test_table1.xlsx'), with_formatting=False)
        self.sheet = self.wbk['Sheet1']


class TestFlexibleRange(unittest.TestCase):
    def setUp(self):
        load_backend('sheetparser.backends._openpyxl')
        self.wbk = load_workbook(os.path.join(os.path.dirname(__file__), 'test_table1.xlsx'), with_formatting=True)
        self.sheet = self.wbk['Sheet1']

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

    def test_flexible(self):
        def has_nocolor(line, linecount):
            return not line[0].is_filled
        pattern = Sheet('result', Columns,
                        FlexibleRange('yellow',
                                      Rows, Line,
                                      Table(table_args=[GetValue, HeaderTableTransform(0, 1),
                                                        FillData]),
                                      stop= has_nocolor
                                      ))
        context = ListContext()
        pattern.match_range(self.wbk['Sheet5'], context)
        dct = dict(context.root)
        self.assertEquals(dct['table'][0].data,
                          [[1, 2, 3], [3, 4, 5]])


class TestComplex(unittest.TestCase):
    def test_complex(self):
        load_backend('sheetparser.backends._openpyxl')
        filename = os.path.join(os.path.dirname(__file__), 'test_table1.xlsx')
        print(filename)
        wbk = load_workbook(filename, with_formatting=True)
        sheet = wbk['Sheet6']
        pattern = Sheet('sheet', Columns,
                        Many(Empty),
                        FlexibleRange('f1',Rows,
                                      Many(Empty),
                                      Table('t1',[GetValue, HeaderTableTransform(2,1),FillData,RemoveEmptyLines('columns')],
                                            stop=no_horizontal), 
                                      Empty, 
                                      FlexibleRange('f2',Columns,
                                                    Many(Empty),Table('t2'),
                                                    stop=no_horizontal),
                                      Many(Empty),
                                      Many((Line('line2',[get_value,Match('Result:')]) 
                                            + Line('line3',[StripLine(),get_value]))
                                           | Line('line1')),
                                      stop = lambda line,linecount: linecount>2 and empty_line(line)
                                      ),
                        Many(Empty),
                        FlexibleRange('f3',Rows,
                                      Many(Empty),
                                      Table('t3',stop = no_horizontal)))
        context = ListContext()
        pattern.match_range(sheet, context)
        dct= context.root
        self.assertEquals(dct.keys(),{'__meta','t1','line2','line1','line3','t1','t2','t3'})
        self.assertEquals(dct['line3'],[['End']])
        self.assertEquals(dct['t1'][0].top_left[0][0],'A more complex example')
        self.assertEquals(dct['t2'][0].top_left,[['Yet another table']])
        self.assertEquals(dct['t3'][0].top_left,[['Another table']])
        
