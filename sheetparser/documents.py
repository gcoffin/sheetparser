# coding: utf-8

"""Definition of the actual items used by the patterns: sheet,
workbook, ranges, lines and columns.  Concrete implementations must be
supplied by backends.
"""

import abc
import importlib
import os
import sys
from abc import abstractmethod

import six

from .utils import ConfigurationError, deprecated


# Documents


class Document(object):
    pass


class RollbackIterator(abc.ABC):
    """An iterator that can save its status to rollback if failure"""

    class SaveStatus(object):
        def __init__(self, rbiter, status, reraise=True):
            self.status = status
            self.rbiter = rbiter
            self.reraise = reraise

        def __enter__(self):
            pass

        def __exit__(self, etype, evalue, etb):
            if etype is not None:
                self.rbiter.idx = self.status
            return not self.reraise  # exception is swallowed if returns true

    def __init__(self, rge):
        self.idx = 0
        self.rge = rge

    def rollback_if_fail(self, reraise=True):
        return RollbackIterator.SaveStatus(self, self.idx, reraise)

    def __iter__(self):
        return self

    def __next__(self):
        if self.is_complete:
            raise StopIteration
        result = self.peek
        self.idx += 1
        return result

    def __repr__(self):
        return "<%s %s %s>" % (
            self.__class__.__name__,
            self.rge, self.idx)

    @property
    @deprecated(' Use is_complete instead')
    def empty(self):
        return self.is_complete


class RbRowIterator(RollbackIterator):
    """Iterates on the rows of a range"""

    @property
    def is_complete(self):
        return self.idx >= self.rge.bottom

    @property
    def peek(self):
        return CellRow(self.rge, self.idx)


class RbVisibleRowIterator(RbRowIterator):
    def __next__(self):
        while True:
            if self.is_complete:
                raise StopIteration
            result = self.peek
            self.idx += 1
            if not result.is_hidden():
                return result


class RbColIterator(RollbackIterator):
    """Iterates on the columns of a range"""

    @property
    def is_complete(self):
        return self.idx >= self.rge.right

    @property
    def peek(self):
        return CellColumn(self.rge, self.idx)


class CellRange(Document):
    """A range (a 2D area) of cells, relative to a parent range"""

    def __init__(self, rge, top=None, left=None, bottom=None, right=None):
        self.rge = rge
        self.top = top or 0
        self.bottom = bottom or rge.height
        self.right = right or rge.width
        self.left = left or 0

    @property
    def width(self):
        return self.right - self.left

    @property
    def height(self):
        return self.bottom - self.top

    def get_range(self):
        return self.top, self.left, self.bottom, self.right

    def cell(self, row, col):
        return self.rge.cell(self.top + row, self.left + col)

    def __repr__(self):
        return "<CellRange %s %s>" % (
            self.rge, (self.top, self.left, self.bottom, self.right))

    def rows(self):
        for row in range(0, self.height):
            yield CellRow(self, row)

    def is_hidden_row(self, row):
        return self.rge.is_hidden_row(row + self.top)

    @property
    def name(self):
        return self.rge.name

    def __str__(self):
        return 'CellRange:<%s>' % (str(list(str(row) for row in self.rows())))


def _abs_index(rge, i):
    if i < 0:
        i = len(rge) + i
        if i <= 0:
            raise IndexError
    elif i >= len(rge):
        raise IndexError
    return i


class CellColumn(CellRange):
    """a vertical line of cells - a range of width 1"""

    def __init__(self, rge, col, top=None, bottom=None):
        self.rge = rge
        self.col = col
        self.top = top or 0
        self.bottom = bottom or rge.height

    @property
    def left(self):
        return self.col

    @property
    def right(self):
        return self.col + 1

    def __getitem__(self, i):
        if isinstance(i, slice):
            return [self[j] for j in list(range(0, len(self)))[i]]
        return self.rge.cell(self.top +
                             _abs_index(self, i), self.col)

    def __len__(self):
        return self.bottom - self.top

    def __repr__(self):
        return "<CellColumn %s %s>" % (self.rge, self.col)

    def __str__(self):
        return 'CellColumn:' + str([i.value for i in self])


class CellRow(CellRange):
    def __init__(self, rge, row, left=None, right=None):
        self.rge = rge
        self._row = row
        self.left = left or 0
        self.right = right or rge.width

    @property
    def top(self):
        return self._row

    @property
    def bottom(self):
        return self._row + 1

    def is_hidden(self):
        return self.rge.is_hidden_row(self._row)

    def __getitem__(self, i):
        if isinstance(i, slice):
            return [self[j] for j in list(range(0, len(self)))[i]]
        return self.rge.cell(self._row,
                             self.left + _abs_index(self, i))

    def __len__(self):
        return self.right - self.left

    def __repr__(self):
        return "<CellRow %s %s>" % (self.rge, self._row)

    def __str__(self):
        return 'CellRow:' + str([i.value for i in self])


BORDER_TOP, BORDER_LEFT, BORDER_BOTTOM, BORDER_RIGHT = (1 << i for i in range(4))
BORDERS_VERTICAL = BORDER_RIGHT | BORDER_LEFT
BORDERS_HORIZONTAL = BORDER_TOP | BORDER_BOTTOM


class SheetDocument(Document, metaclass=abc.ABCMeta):
    """Base class for sheets, to be implemented
    by a backend"""

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @abstractmethod
    def cell(self, row, col):
        raise NotImplementedError

    @abstractmethod
    def is_hidden(self):
        raise NotImplementedError


class WorkbookDocument(Document, metaclass=abc.ABCMeta):
    pass


def load_backend(name, ignore_fail=False):
    try:
        return importlib.import_module(name)
    except ImportError:
        print("Coudn't import file %s" % name, file=sys.stderr)
        if not ignore_fail:
            raise
        return None


class LazyModule(object):
    def __init__(self, name):
        self.name = name
        self.module = None

    def __getattr__(self, attr):
        if self.module is None:
            self.module = load_backend(self.name)
        return getattr(self.module, attr)


class WorkbookReader(dict):
    """a callable object that will call the proper
    backend to read the file"""

    def __init__(self):
        _openpyxl = LazyModule('sheetparser.backends._openpyxl')
        _xlrd = LazyModule('sheetparser.backends._xlrd')
        _pdfminer = LazyModule('sheetparser.backends._pdfminer')
        self['.xls', True] = _xlrd
        self['.xlsx', False] = _xlrd
        self['.xlsm', False] = _xlrd
        self['_xlrd'] = _xlrd

        self['.xlsx', True] = _openpyxl
        self['.xlsm', True] = _openpyxl
        self['.pdf', False] = _pdfminer
        self['_openpyxl'] = _openpyxl

    def __call__(self, filepath, with_formatting=False, with_backend=None):
        backend = None
        if with_backend is None:
            __, ext = os.path.splitext(filepath)
            backend = self.get((ext, with_formatting), None)
        elif with_backend:
            if with_backend in self:
                backend = self[with_backend]
            else:
                backend = self[with_backend] = load_backend(with_backend)
        if not backend:
            raise ConfigurationError(
                "You need to import a backend that"
                " provides this functionality first")
        else:
            return backend.load_workbook(filepath,
                                         with_formatting=with_formatting)


load_workbook = WorkbookReader()
