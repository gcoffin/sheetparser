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
        self.bottom, self.right = self.rge.shape

    def is_hidden(self):
        return False

    def abscell(self, row, col):
        return DummyCell(self.rge[row, col])

    cell = abscell  # NO!!

    def __repr__(self):
        return "<DummySheet %s>" % self.rge


class DummyCell(object):
    def __init__(self, value):
        self.value = value
        self.is_merged = False

    @property
    def is_empty(self):
        if isinstance(self.value,np.ndarray):
            return np.isnan(self.value[0])
        return self.value == ''


def to_list_value(l):
    return [i.value for i in l]


class TestColIterator(unittest.TestCase):
    def test_row_iters(self):
        test_array = [[0, 0, 1, 1, 0]] * 3
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
        r = CellRange(sheet, 0, 0, 1, 1)  # top left is 7, bottom right is 13
        it = RbColIterator(r)
        sub_array = test_array[0:1, 0:1]
        for col in zip(*sub_array):
            self.assertSequenceEqual(to_list_value(six.next(it)), list(col))

    #  0  1  2  3  4
    #  5  6  7  8  9
    #  10 11 12 13 14
    #  15 16 17 18 19
    def test_subrange2(self):
        test_array = np.arange(20, dtype=int).reshape(4, 5)
        sheet = DummySheet('test', test_array)
        r = CellRange(sheet, 1, 2, 3, 4)  # top left is 7, bottom right is 13
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
        r = CellRange(sheet, 1, 1, 5, 5)  # top left is 7, bottom right is 13
        sr = CellRange(r, 1, 1, 2, 2)
        it = RbRowIterator(sr)
        for row in test_array[2:3, 2:3]:
            self.assertSequenceEqual(to_list_value(six.next(it)), list(row))


class TestArray(unittest.TestCase):
    def test_rollback(self):
        test_array = np.array([[1]*5])
        sheet = DummySheet('test', test_array)
        pattern = Sheet('result', Rows,
                        Many(Line, min=2) | Line('line'))
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertSequenceEqual(context.root['or_']['line'], [1, 1, 1, 1, 1])


class TestBug(unittest.TestCase):
    def test_many_many(self):
        sheet = DummySheet('dummy', [['h']*2, ['l', 'd'], ['']*2])
        pattern = Sheet('e', Rows,
                        Many('tables',
                             Sequence(Table('table',
                                            table_args=[GetValue, HeaderTableTransform(1, 1), FillData,
                                                        TableNotEmpty],
                                            stop=empty_line),
                                      Many('between tables2', Empty))))
        context = PythonObjectContext()
        pattern.match_range(sheet, context)
        self.assertEqual(len(context.tables), 1)
