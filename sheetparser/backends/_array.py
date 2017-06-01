"""uses openpyxl as the engine to read the Excel sheet."""

from __future__ import print_function
import csv
import os
import six
import logging

from ..documents import (CellRange, SheetDocument, WorkbookDocument,
                         load_workbook)
from ..utils import EMPTY_CELL

logger = logging.getLogger('sheetparser')

class rawCell(object):
    def __init__(self, row, column,value):
        self.value = EMPTY_CELL if value is None else value
        self.row = row
        self.column = column

    @property
    def is_empty(self):
        return self.value == EMPTY_CELL

    @property
    def is_merged(self):
        return False

class rawSheet(CellRange, SheetDocument):
    def __init__(self, name, values):
        self.name = name
        self.data = values
        self.top, self.left = (0, 0)
        self.bottom = len(values)
        self.right = max(len(i) for i in values)

    def is_hidden(self):
        return False

    def cell(self, row, col):
        try:
            return rawCell(row, col, self.data[row][col])
        except IndexError:
            return rawCell(row,col,None)

    def __repr__(self):
        return "<rawSheet %s>" % self.name

class rawWorkbook(WorkbookDocument):
    """A class to open workbooks and obtain sheets"""
    def __init__(self, values_map):
        self.data = values_map
        
    def __iter__(self):
        return (rawSheet(i,data) for i,data in six.iteritems(self.data))

    def __getitem__(self, name_or_id):
        return rawSheet(name_or_id,self.data[name_or_id])

def load_workbook(filename,options=None):
    options = options or {}
    with open(filename,'rb') as f:
        return rawWorkbook({ 1 : list(csv.reader(f,**options)) })
